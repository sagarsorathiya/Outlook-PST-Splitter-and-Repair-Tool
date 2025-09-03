"""Core splitting logic coordinating OutlookSession and size grouping."""
import logging
import traceback
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from threading import Event
from typing import Callable, Iterable, Optional, Any

from .outlook import MailItemInfo, OutlookSession, PSTSizeExceededException, PSTRecoveryManager
from .space_crisis_manager import AdvancedSpaceManager
from .ultimate_pst_handler import UltimatePSTHandler
from .log_exporter import log_exporter
from .performance_optimizer import HighPerformancePSTProcessor
from .util import format_bytes


@dataclass
class SplitResult:
    """Summary of splitting operation."""

    created_files: list[Path]
    total_items: int
    total_bytes: int = 0
    errors: list[str] | None = None


def group_items_by_size(items: Iterable[MailItemInfo], max_bytes: int) -> list[list[MailItemInfo]]:
    """Greedy size-based grouping of items into buckets.

    Each bucket's summed `size` will not exceed max_bytes, unless a single item
    itself exceeds max_bytes in which case that item forms its own bucket.
    """
    buckets: list[list[MailItemInfo]] = []
    current: list[MailItemInfo] = []
    current_size = 0
    for it in items:
        size = getattr(it, "size", 0)
        if size > max_bytes:  # oversize single item
            if current:
                buckets.append(current)
                current = []
                current_size = 0
            buckets.append([it])
            continue
        if current_size + size > max_bytes and current:
            buckets.append(current)
            current = [it]
            current_size = size
        else:
            current.append(it)
            current_size += size
    if current:
        buckets.append(current)
    return buckets


def group_items_by_month(items: Iterable[MailItemInfo]) -> dict[str, list[MailItemInfo]]:
    """Groups items by YYYY-MM format."""
    groups: dict[str, list[MailItemInfo]] = defaultdict(list)
    for item in items:
        month_key = (item.received or datetime.min).strftime("%Y-%m")
        groups[month_key].append(item)
    return dict(groups)


def group_items_by_year(items: Iterable[MailItemInfo]) -> dict[str, list[MailItemInfo]]:
    """Groups items by YYYY format with strict year separation."""
    groups: dict[str, list[MailItemInfo]] = defaultdict(list)
    unknown_items = []
    
    for item in items:
        try:
            if item.received:
                year_key = item.received.strftime("%Y")
                groups[year_key].append(item)
            else:
                # Handle items without received date
                unknown_items.append(item)
        except (AttributeError, ValueError) as e:
            logging.debug(f"Failed to extract year from item: {e}")
            unknown_items.append(item)
    
    # Add unknown items as separate group if any exist
    if unknown_items:
        groups["Unknown_Date"] = unknown_items
        logging.info(f"📅 Found {len(unknown_items)} items with unknown dates")
    
    # Log year distribution for analysis
    for year, year_items in groups.items():
        log_exporter.log_group_creation(
            f"Year_{year}", 
            len(year_items), 
            sum(getattr(item, 'size', 0) for item in year_items)
        )
        logging.info(f"📅 Year {year}: {len(year_items)} items")
    
    return dict(groups)


def _group_by_folder(items: Iterable[MailItemInfo]) -> dict[str, list[MailItemInfo]]:
    """Group items by their top-level folder path."""
    groups: dict[str, list[MailItemInfo]] = defaultdict(list)
    for item in items:
        folder_path = item.folder_path if item.folder_path else ""
        # Get top-level folder name
        if folder_path:
            top_folder = folder_path.split('/')[0]
        else:
            top_folder = "Root"
        groups[top_folder].append(item)
    return dict(groups)


def check_pst_health(source_pst: Path) -> dict[str, Any]:
    """Check PST file health and available space before splitting.
    
    Returns:
        dict: Health status with size info, recommendations, and warnings
    """
    health_report = {
        "healthy": True,
        "warnings": [],
        "size_info": {},
        "recommendations": []
    }
    
    try:
        # Use advanced space manager for comprehensive analysis
        space_manager = AdvancedSpaceManager()
        crisis_analysis = space_manager.analyze_pst_space_crisis(source_pst)
        
        # Convert space crisis analysis to health report format
        health_report["size_info"] = {
            "current_size": source_pst.stat().st_size,
            "current_size_formatted": f"{crisis_analysis['pst_size_gb']:.2f} GB",
            "pst_type": "Unicode" if crisis_analysis['pst_size_gb'] > 2 else "ANSI",
            "utilization_percent": crisis_analysis['size_percentage'],
            "estimated_free_space": crisis_analysis.get('free_space_bytes', 0),
            "free_space_formatted": f"{crisis_analysis.get('free_space_gb', 0):.2f} GB"
        }
        
        # Convert risk levels to health status
        if crisis_analysis['risk_level'] == 'CRITICAL':
            health_report["healthy"] = False
            health_report["warnings"].append(f"CRITICAL: PST is {crisis_analysis['size_percentage']:.1f}% full - high risk of failures")
        elif crisis_analysis['risk_level'] == 'HIGH':
            health_report["warnings"].append(f"HIGH RISK: PST is {crisis_analysis['size_percentage']:.1f}% full - may encounter issues")
        elif crisis_analysis['risk_level'] == 'MEDIUM':
            health_report["warnings"].append(f"MEDIUM RISK: PST is {crisis_analysis['size_percentage']:.1f}% full - monitor closely")
            
        # Add space issues as warnings
        if crisis_analysis['space_issues']:
            health_report["warnings"].extend(crisis_analysis['space_issues'])
            
        # Add recommendations
        if crisis_analysis['recommendations']:
            health_report["recommendations"].extend(crisis_analysis['recommendations'])
            
        # Basic file checks
        if not source_pst.exists():
            health_report["healthy"] = False
            health_report["warnings"].append(f"PST file not found: {source_pst}")
            return health_report
            
        # Check if file is locked/in use
        try:
            with open(source_pst, 'r+b') as f:
                pass
        except PermissionError:
            health_report["warnings"].append("PST file may be locked by Outlook - ensure Outlook is closed")
            
    except Exception as e:
        health_report["healthy"] = False
        health_report["warnings"].append(f"Could not analyze PST health: {e}")
        
    return health_report


def split_pst(
    source_pst: Path,
    output_dir: Path,
    mode: str,
    max_size_bytes: int | None,
    cancel: Event,
    progress_cb: Callable[[int, int, int, int], None] | None = None,
    dry_run: bool = False,
    include_non_mail: bool = False,
    move_items: bool = False,
    verify: bool = True,
    fast_enumeration: bool = False,
    turbo_mode: bool = False,  # Extreme performance mode with space safety
    suppress_item_logs: bool = False,
    stream_size_mode: bool = False,
    throttle_progress_ms: int = 250,
    include_folders: Optional[set[str]] = None,
    exclude_folders: Optional[set[str]] = None,
    sender_domains: Optional[set[str]] = None,
    date_range: Optional[tuple[Optional[datetime], Optional[datetime]]] = None,
    summary_csv: Optional[Path] = None,
) -> SplitResult:
    """Split PST according to mode with advanced space crisis management.

    mode: 'size' | 'year' | 'month'
    max_size_bytes: required when mode == 'size'
    """
    if mode == "size" and (not max_size_bytes or max_size_bytes <= 0):
        raise ValueError("max_size_bytes must be positive for size mode")

    # Initialize session logging
    log_exporter.log_session_start(
        str(source_pst), str(output_dir), mode,
        max_size_bytes=max_size_bytes,
        include_folders=list(include_folders) if include_folders else None,
        exclude_folders=list(exclude_folders) if exclude_folders else None,
        sender_domains=list(sender_domains) if sender_domains else None,
        turbo_mode=turbo_mode,
        move_items=move_items
    )

    logging.info(
        "Starting split with advanced space management: source=%s mode=%s %s%s%s%s%s",
        source_pst,
        mode,
        f"limit={format_bytes(max_size_bytes)} " if (mode == 'size' and max_size_bytes) else "",
        f" include_folders={sorted(include_folders)}" if include_folders else "",
        f" exclude_folders={sorted(exclude_folders)}" if exclude_folders else "",
        f" sender_domains={sorted(sender_domains)}" if sender_domains else "",
        f" date_range={[dr.isoformat() if dr else None for dr in date_range]}" if date_range else "",
    )

    # Initialize advanced space manager
    space_manager = AdvancedSpaceManager()
    
    # Initialize Ultimate PST Handler for infinite loop scenarios
    ultimate_handler = UltimatePSTHandler()
    
    # Track space liberation attempts to detect infinite loops
    space_liberation_attempts = []
    max_liberation_attempts = 3  # Allow up to 3 attempts before switching to ultimate handler
    
    # Perform comprehensive PST health check
    try:
        logging.info("🔍 Running comprehensive PST space analysis...")
        crisis_analysis = space_manager.analyze_pst_space_crisis(source_pst)
        
        # Log space analysis results
        logging.info(f"📊 PST Analysis: {crisis_analysis['pst_size_gb']:.2f} GB, Risk: {crisis_analysis['risk_level']}")
        
        # Handle critical space situations
        if crisis_analysis['risk_level'] == 'CRITICAL':
            logging.warning("⚠️ CRITICAL space situation detected!")
            try:
                # Create crisis management plan
                crisis_plan = space_manager.create_space_crisis_plan(source_pst)
                if crisis_plan['requires_manual_intervention']:
                    logging.error("Manual intervention required for space crisis resolution")
                    # Continue anyway with maximum safety measures
                    turbo_mode = False  # Disable turbo mode for safety
                    logging.info("🛡️ Disabled turbo mode for safety in critical space situation")
                    
            except Exception as e:
                logging.warning(f"Space crisis planning failed: {e}")
    
    except Exception as e:
        logging.warning(f"Advanced space analysis failed, using basic health check: {e}")
        
    # Perform basic PST health check
    health_report = check_pst_health(source_pst)
    if health_report["warnings"]:
        for warning in health_report["warnings"]:
            logging.warning("PST Health Check: %s", warning)
    if health_report["recommendations"]:
        for rec in health_report["recommendations"]:
            logging.info("PST Health Recommendation: %s", rec)
    
    if not health_report["healthy"]:
        logging.error("PST health check failed. Proceeding with enhanced safety measures...")
        turbo_mode = False  # Disable turbo mode for safety
    else:
        logging.info("PST health check passed - %s PST, %s used (%s free)", 
                    health_report["size_info"]["pst_type"],
                    health_report["size_info"]["current_size_formatted"],
                    health_report["size_info"]["free_space_formatted"])

    # Normalise folder filters to lower for case-insensitive match on top-level folder
    if include_folders:
        include_folders = {f.strip().lower() for f in include_folders if f.strip()}
    if exclude_folders:
        exclude_folders = {f.strip().lower() for f in exclude_folders if f.strip()}
    if sender_domains:
        sender_domains = {d.strip().lower() for d in sender_domains if d.strip()}

    created_files: list[Path] = []
    errors: list[str] = []
    total_items = 0
    total_bytes = 0
    outlook = None

    try:
        # Initialize Outlook session with recovery manager
        outlook = OutlookSession()
        recovery_manager = PSTRecoveryManager(outlook)
        
        # Initialize performance optimizer
        performance_processor = HighPerformancePSTProcessor(outlook)
        if turbo_mode:
            performance_processor.set_performance_mode(True, batch_size=100)  # Larger batches in turbo mode
            log_exporter.log_performance_metric("turbo_mode", "enabled", "mode")
        else:
            performance_processor.set_performance_mode(True, batch_size=50)   # Standard batches
        
        # Attach source PST
        outlook.attach_pst(source_pst)
        source_store_name = outlook.find_store_by_path(source_pst)
        if not source_store_name:
            raise RuntimeError(f"Could not find attached store for {source_pst}")

        # Enumerate items with space safety
        logging.info("📧 Enumerating items...")
        try:
            items = list(outlook.iter_mail_items(
                source_store_name,
                include_non_mail=include_non_mail,
                cancel_event=cancel
            ))
        except PSTSizeExceededException as e:
            logging.error(f"PST size limit exceeded during enumeration: {e}")
            # Attempt recovery with space manager
            try:
                logging.info("🔄 Attempting enumeration recovery...")
                items = list(outlook.iter_mail_items(
                    source_store_name,
                    include_non_mail=include_non_mail,
                    cancel_event=cancel
                ))
                logging.info("✅ Enumeration recovery successful")
            except Exception as recovery_error:
                logging.error(f"Enumeration recovery failed: {recovery_error}")
                raise

        if cancel.is_set():
            logging.info("Operation cancelled during enumeration")
            return SplitResult(created_files, 0, 0, ["Operation cancelled"])

        # Apply filters
        if include_folders or exclude_folders or sender_domains or date_range:
            original_count = len(items)
            items = _apply_filters(items, include_folders, exclude_folders, sender_domains, date_range)
            logging.info("Filtered %d → %d items", original_count, len(items))

        total_items = len(items)
        logging.info("Found %d items to process", total_items)

        if total_items == 0:
            logging.warning("No items found to split")
            return SplitResult([], 0, 0)

        # Group items according to mode
        if mode == "size":
            if max_size_bytes is None:
                raise ValueError("max_size_bytes cannot be None for size mode")
            groups = group_items_by_size(items, max_size_bytes)
            group_names = [f"part{i+1:03d}" for i in range(len(groups))]
        elif mode == "year":
            grouped_dict = group_items_by_year(items)
            groups = list(grouped_dict.values())
            group_names = list(grouped_dict.keys())
        elif mode == "month":
            grouped_dict = group_items_by_month(items)
            groups = list(grouped_dict.values())
            group_names = list(grouped_dict.keys())
        else:
            raise ValueError(f"Unknown mode: {mode}")

        logging.info("Split into %d groups: %s", len(groups), group_names)

        # Track cumulative progress across all groups
        cumulative_processed = 0
        total_bytes_processed = 0
        
        # Process each group with enhanced error handling
        for i, (group_items, group_name) in enumerate(zip(groups, group_names)):
            if cancel.is_set():
                logging.info("Operation cancelled during processing")
                break

            if not group_items:
                continue

            group_item_count = len(group_items)
            logging.info("Processing group %d/%d: %s (%d items)", i+1, len(groups), group_name, group_item_count)

            # Create target PST with space monitoring
            target_filename = f"{source_pst.stem}_{group_name}.pst"
            target_path = output_dir / target_filename
            target_store_name = None

            try:
                if not dry_run:
                    outlook.create_new_pst(target_path)
                    target_store_name = outlook.find_store_by_path(target_path)
                    if not target_store_name:
                        raise RuntimeError(f"Could not find created store for {target_path}")

                # Use high-performance batch copying instead of individual items
                logging.info(f"📦 Processing {group_item_count} items using optimized batch processing...")
                
                if not dry_run and target_store_name:
                    # Create a progress callback that reports cumulative progress
                    def copy_progress_callback(msg):
                        if progress_cb:
                            # Extract current progress from message if possible
                            try:
                                if "Processed" in msg and "/" in msg:
                                    # Parse "Processed X/Y items" format
                                    parts = msg.split()
                                    if len(parts) >= 2:
                                        progress_part = parts[1]  # "X/Y"
                                        current_in_batch = int(progress_part.split('/')[0])
                                        current_total = cumulative_processed + current_in_batch
                                        progress_cb(current_total, total_items, 0, 0)
                            except:
                                # Fallback to basic group progress
                                progress_cb(cumulative_processed, total_items, 0, 0)
                    
                    # Use the performance optimizer for batch processing
                    copy_results = performance_processor.copy_items_optimized(
                        group_items,
                        target_store_name,
                        "Items",
                        target_path,
                        move_items=move_items,
                        progress_callback=copy_progress_callback,
                        cancel_event=cancel
                    )
                    
                    # Check if operation was cancelled during copy
                    if copy_results.get('cancelled', False):
                        logging.info("Operation cancelled during copy/move")
                        break
                    
                    processed_in_group = copy_results['success_count']
                    
                    # Update cumulative progress
                    cumulative_processed += processed_in_group
                    
                    # Report final progress for this group
                    if progress_cb:
                        progress_cb(cumulative_processed, total_items, 0, 0)
                    
                    # Log any failed items
                    if copy_results['failed_items']:
                        logging.warning(f"⚠️ {len(copy_results['failed_items'])} items failed to process")
                        for failed_item in copy_results['failed_items']:
                            log_exporter.log_error(
                                "item_copy_failed", 
                                failed_item['error'], 
                                {'item_id': failed_item['item_id'], 'year': failed_item['year']}
                            )
                    
                    # Log performance metrics
                    if copy_results['total_time_ms'] > 0:
                        items_per_second = processed_in_group / (copy_results['total_time_ms'] / 1000)
                        logging.info(f"⚡ Performance: {items_per_second:.1f} items/sec, {copy_results['avg_time_per_item_ms']:.1f}ms per item")
                else:
                    # Dry run - just count items
                    processed_in_group = group_item_count
                    cumulative_processed += processed_in_group
                    
                    # Report progress for dry run
                    if progress_cb:
                        progress_cb(cumulative_processed, total_items, 0, 0)

                # Calculate group size for reporting
                group_bytes = sum(getattr(item, "size", 0) for item in group_items)
                total_bytes += group_bytes
                
                if not dry_run and target_path.exists():
                    created_files.append(target_path)
                    logging.info("✅ Created %s with %d items (%s)", target_filename, processed_in_group, format_bytes(group_bytes))

            except Exception as group_error:
                error_msg = f"Failed to process group {group_name}: {group_error}"
                logging.error(error_msg)
                errors.append(error_msg)

    except Exception as e:
        error_msg = f"Split operation failed: {e}"
        logging.error(error_msg)
        logging.error(traceback.format_exc())
        errors.append(error_msg)

    finally:
        # Cleanup
        try:
            if 'outlook' in locals() and outlook:
                if source_pst:
                    try:
                        outlook.detach_pst(source_pst)
                    except Exception:
                        pass  # Ignore detach errors
                for created_file in created_files:
                    try:
                        outlook.detach_pst(created_file)
                    except Exception:
                        pass  # Ignore detach errors
        except Exception as e:
            logging.warning(f"Cleanup error: {e}")

    logging.info("Split completed: %d files created, %d total items, %s total size", 
                len(created_files), total_items, format_bytes(total_bytes))
    
    # Export analysis report
    try:
        analysis_dir = log_exporter.export_analysis_report(output_dir)
        logging.info(f"📊 Analysis report exported to: {analysis_dir}")
    except Exception as e:
        logging.warning(f"Failed to export analysis report: {e}")
    
    return SplitResult(created_files, total_items, total_bytes, errors if errors else None)


def _apply_filters(
    items: list[MailItemInfo],
    include_folders: Optional[set[str]],
    exclude_folders: Optional[set[str]],
    sender_domains: Optional[set[str]],
    date_range: Optional[tuple[Optional[datetime], Optional[datetime]]],
) -> list[MailItemInfo]:
    """Apply various filters to the item list."""
    filtered = items

    # Folder filters
    if include_folders:
        filtered = [item for item in filtered if _matches_folder_filter(item.folder_path, include_folders)]
    if exclude_folders:
        filtered = [item for item in filtered if not _matches_folder_filter(item.folder_path, exclude_folders)]

    # Sender domain filter
    if sender_domains:
        filtered = [item for item in filtered if _matches_sender_domain(item.sender_email or "", sender_domains)]

    # Date range filter
    if date_range:
        start_date, end_date = date_range
        if start_date or end_date:
            filtered = [item for item in filtered if _matches_date_range(item.received, start_date, end_date)]

    return filtered


def _matches_folder_filter(folder_path: str, folder_filters: set[str]) -> bool:
    """Check if folder path matches any of the folder filters."""
    if not folder_path:
        return False
    
    # Extract top-level folder name and normalize
    folder_parts = folder_path.split("\\")
    if len(folder_parts) >= 2:  # Skip root store name
        top_folder = folder_parts[1].lower()
        return top_folder in folder_filters
    return False


def _matches_sender_domain(sender_email: str, sender_domains: set[str]) -> bool:
    """Check if sender email domain matches any of the domain filters."""
    if not sender_email or "@" not in sender_email:
        return False
    
    domain = sender_email.split("@")[-1].lower()
    return domain in sender_domains


def _matches_date_range(
    item_date: Optional[datetime],
    start_date: Optional[datetime],
    end_date: Optional[datetime]
) -> bool:
    """Check if item date falls within the specified range."""
    if item_date is None:
        return False
    if start_date and item_date < start_date:
        return False
    if end_date and item_date > end_date:
        return False
    return True
