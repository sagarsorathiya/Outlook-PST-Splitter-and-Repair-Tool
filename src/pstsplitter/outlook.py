"""Outlook COM wrapper abstractions.

This module isolates pywin32 / Outlook specific operations so the rest of the
code base can remain more testable. Only Windows with Outlook installed is
supported. All functions raise RuntimeError if Outlook automation is not
available.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Optional, List, Dict, Tuple, Any
from datetime import datetime
from threading import Event
import logging
import time

try:  # pragma: no cover - optional dependency, ignore if not installed
    import win32com.client  # type: ignore[import]
    from win32com.client import Dispatch  # type: ignore[import]
    # Try to import win32timezone to avoid runtime errors
    try:
        import win32timezone  # type: ignore[import]
    except ImportError:
        # Create a dummy win32timezone to prevent import errors
        import sys
        import types
        win32timezone = types.ModuleType('win32timezone')
        sys.modules['win32timezone'] = win32timezone
except Exception:  # pragma: no cover
    win32com = None  # type: ignore
    Dispatch = None  # type: ignore


from .space_crisis_manager import AdvancedSpaceManager

class PSTSizeExceededException(Exception):
    """Raised when a PST file has reached its maximum size limit."""
    pass


class PSTRecoveryManager:
    """Manages PST size issues and recovery strategies."""
    
    def __init__(self, outlook_session):
        self.session = outlook_session
        self.temp_cleanup_performed = False
        self.consecutive_failures = 0
        self.advanced_manager = AdvancedSpaceManager()
    
    def get_pst_free_space(self, pst_path: Path) -> tuple[int, int]:
        """Get PST file size and estimated free space.
        
        Returns:
            tuple: (current_size_bytes, estimated_free_bytes)
        """
        try:
            if pst_path.exists():
                current_size = pst_path.stat().st_size
                # Estimate based on known PST limits
                if current_size < 2_000_000_000:  # ~2GB (ANSI PST)
                    max_size = 2_147_483_648  # 2GB
                else:  # Unicode PST
                    max_size = 53_687_091_200  # ~50GB
                
                free_space = max(0, max_size - current_size)
                return current_size, free_space
            return 0, 0
        except Exception as e:
            logging.debug("Could not determine PST size for %s: %s", pst_path, e)
            return 0, 0
    
    def attempt_pst_cleanup(self, pst_path: Path) -> bool:
        """Attempt to free space in PST by cleaning deleted items.
        
        Returns:
            bool: True if cleanup was attempted
        """
        if self.temp_cleanup_performed:
            return False
            
        try:
            logging.info("Attempting to free space in source PST...")
            # Try to find and empty deleted items folder
            if hasattr(self.session, 'app') and self.session.app:
                for store in self.session.app.Session.Stores:
                    store_path = getattr(store, 'FilePath', '')
                    if store_path and Path(store_path).resolve() == pst_path.resolve():
                        try:
                            # Find deleted items folder
                            deleted_items = None
                            for folder in store.GetRootFolder().Folders:
                                if getattr(folder, 'DefaultItemType', 0) == 3:  # olMailItem
                                    if 'deleted' in getattr(folder, 'Name', '').lower():
                                        deleted_items = folder
                                        break
                            
                            if deleted_items and deleted_items.Items.Count > 0:
                                logging.info("Found %d items in Deleted Items, attempting cleanup...", 
                                           deleted_items.Items.Count)
                                # Strategy 1: Move items to a temporary subfolder to free immediate space
                                temp_folder_name = f"TempCleanup_{int(time.time())}"
                                try:
                                    temp_folder = deleted_items.Folders.Add(temp_folder_name)
                                    moved_count = 0
                                    # Move items in batches to avoid timeout
                                    for item in list(deleted_items.Items)[:100]:  # Limit to 100 items
                                        try:
                                            item.Move(temp_folder)
                                            moved_count += 1
                                        except:
                                            break
                                    if moved_count > 0:
                                        logging.info("Moved %d items to temporary cleanup folder", moved_count)
                                        self.temp_cleanup_performed = True
                                        return True
                                except Exception as e:
                                    logging.debug("Temp folder cleanup failed: %s", e)
                            else:
                                # Strategy 2: Try to force PST compaction
                                logging.info("No deleted items found, attempting PST optimization...")
                                return self._attempt_pst_optimization(store)
                                    
                        except Exception as e:
                            logging.debug("Could not access store folders: %s", e)
                            
        except Exception as e:
            logging.debug("PST cleanup attempt failed: %s", e)
        
        return False
    
    def _attempt_pst_optimization(self, store) -> bool:
        """Attempt PST optimization strategies."""
        try:
            # Strategy 1: Force save all items to commit changes
            logging.info("Attempting PST optimization by saving pending changes...")
            
            # Try to access and save any unsaved items
            for folder in store.GetRootFolder().Folders:
                try:
                    if hasattr(folder, 'Items') and folder.Items.Count > 0:
                        # Force save on a few recent items to commit changes
                        for i, item in enumerate(folder.Items):
                            if i >= 5:  # Limit to 5 items to avoid timeout
                                break
                            try:
                                if hasattr(item, 'Save'):
                                    item.Save()
                            except:
                                pass
                except:
                    continue
            
            logging.info("PST optimization attempted")
            return True
            
        except Exception as e:
            logging.debug("PST optimization failed: %s", e)
            return False
    
    def handle_space_crisis(self, pst_path: Path) -> Dict[str, Any]:
        """Advanced space crisis handling with comprehensive analysis."""
        try:
            logging.warning("🚨 SPACE CRISIS DETECTED - Running advanced analysis for %s", pst_path)
            
            # Comprehensive space analysis
            crisis_analysis = self.advanced_manager.analyze_pst_space_crisis(pst_path)
            
            # Log critical findings
            for issue in crisis_analysis.get("critical_issues", []):
                logging.error("🔴 CRITICAL: %s", issue)
            for warning in crisis_analysis.get("warnings", []):
                logging.warning("🟡 WARNING: %s", warning)
                
            # Determine if emergency liberation is needed
            space_details = crisis_analysis.get("space_details", {})
            utilization = space_details.get("utilization_percent", 0)
            free_space_mb = space_details.get("free_space_mb", 0)
            
            crisis_response = {
                "analysis": crisis_analysis,
                "emergency_actions_taken": [],
                "recommendations": crisis_analysis.get("recommendations", []),
                "success": False
            }
            
            # Trigger emergency liberation if needed
            if utilization > 95 or free_space_mb < 100:
                logging.warning("🚨 Triggering emergency space liberation protocol...")
                liberation_result = self.advanced_manager.emergency_space_liberation(pst_path, self.session)
                crisis_response["emergency_actions_taken"] = liberation_result.get("actions_taken", [])
                crisis_response["space_freed_mb"] = liberation_result.get("space_freed_mb", 0)
                crisis_response["success"] = liberation_result.get("success", False)
                
                if crisis_response["success"]:
                    logging.info("✅ Emergency space liberation successful: %.1fMB freed", 
                               crisis_response["space_freed_mb"])
                else:
                    logging.error("❌ Emergency space liberation failed")
                    
            # Create crisis management plan
            crisis_plan = self.advanced_manager.create_space_crisis_plan(pst_path)
            crisis_response["crisis_plan"] = crisis_plan
            
            # Log immediate action recommendations
            immediate_actions = crisis_plan.get("immediate_actions", [])
            if immediate_actions:
                logging.warning("🔧 IMMEDIATE ACTIONS REQUIRED:")
                for action in immediate_actions:
                    logging.warning("   %s", action)
                    
            return crisis_response
            
        except Exception as e:
            logging.exception("Space crisis handling failed")
            return {
                "analysis": {"critical_issues": [f"Crisis handling failed: {e}"]},
                "emergency_actions_taken": [],
                "recommendations": ["Manual PST cleanup required"],
                "success": False
            }


@dataclass
class MailItemInfo:
    """Lightweight representation of an Outlook item we might move/copy.

    received: naive datetime (local) or None if not applicable.
    folder_path: path relative to store root (e.g. "Inbox/SubFolder")
    """

    entry_id: str
    subject: str
    size: int  # bytes
    received: Optional[datetime] = None
    folder_path: str = ""
    sender_email: Optional[str] = None


class OutlookSession:
    """Represents an Outlook MAPI session for PST manipulation."""

    def __init__(self) -> None:
        if Dispatch is None:
            raise RuntimeError("Outlook COM automation unavailable (pywin32 / Outlook missing)")
        self._app = Dispatch("Outlook.Application")
        self._namespace = self._app.GetNamespace("MAPI")
        # Cache for destination folders to reduce COM traversal repetition
        self._folder_cache: Dict[Tuple[str, str], object] = {}
        # When True, suppress per-item copy debug logs for performance on huge PSTs
        self.suppress_item_logs: bool = False
        # Recovery manager for PST size issues
        self.recovery_manager = PSTRecoveryManager(self)

    # --- PST Attachment / Creation -------------------------------------------------
    def attach_pst(self, pst_path: Path) -> None:
        logging.info("Attaching PST: %s", pst_path)
        self._namespace.AddStore(str(pst_path))
        self._folder_cache.clear()
        # Immediately log available stores for diagnostics
        try:
            stores = self._namespace.Stores
            names = []
            for i in range(1, stores.Count + 1):  # type: ignore[attr-defined]
                st = stores.Item(i)
                try:
                    names.append(f"{st.DisplayName} -> {getattr(st, 'FilePath', '?')}")
                except Exception:
                    names.append(st.DisplayName)
            logging.debug("Stores after attach: %s", names)
        except Exception as e:  # pragma: no cover
            logging.debug("Failed listing stores: %s", e)

    def detach_pst(self, pst_path: Path) -> None:
        logging.info("Detaching PST: %s", pst_path)
        stores = self._namespace.Stores
        for i in range(1, stores.Count + 1):
            store = stores.Item(i)
            if Path(store.FilePath).resolve() == pst_path.resolve():
                self._namespace.RemoveStore(store.GetRootFolder())
                break

    def create_new_pst(self, pst_path: Path) -> None:
        if pst_path.exists():
            logging.warning("PST already exists, will attach: %s", pst_path)
        logging.info("Creating new PST: %s", pst_path)
        self._namespace.AddStore(str(pst_path))
        self._folder_cache.clear()
        # Attempt to rename root folder to file stem for uniqueness
        try:
            desired = pst_path.stem
            stores = self._namespace.Stores
            for i in range(1, stores.Count + 1):  # type: ignore[attr-defined]
                store = stores.Item(i)
                try:
                    if Path(store.FilePath).resolve() == pst_path.resolve():
                        root = store.GetRootFolder()
                        if root.Name != desired:
                            root.Name = desired  # type: ignore[attr-defined]
                            logging.debug("Renamed new store root to %s", desired)
                        break
                except Exception:
                    continue
        except Exception as e:  # pragma: no cover
            logging.debug("Store rename failed for %s: %s", pst_path, e)

    def rename_store_by_path(self, pst_path: Path, new_name: str) -> None:
        try:
            stores = self._namespace.Stores
            for i in range(1, stores.Count + 1):  # type: ignore[attr-defined]
                store = stores.Item(i)
                try:
                    if Path(store.FilePath).resolve() == pst_path.resolve():
                        root = store.GetRootFolder()
                        root.Name = new_name  # type: ignore[attr-defined]
                        logging.info("Renamed store %s to %s", pst_path, new_name)
                        return
                except Exception:
                    continue
        except Exception as e:  # pragma: no cover
            logging.debug("Failed renaming store %s: %s", pst_path, e)

    # --- Store Resolution ----------------------------------------------------------
    def find_store_by_path(self, pst_path: Path) -> Optional[str]:
        """Return display name of store that matches pst_path, else None."""
        stores = self._namespace.Stores
        target = pst_path.resolve()
        # First pass: match on FilePath
        for i in range(1, stores.Count + 1):  # type: ignore[attr-defined]
            store = stores.Item(i)
            try:
                fp = Path(store.FilePath).resolve()
                if fp == target:
                    return store.DisplayName
            except Exception:
                continue
        # Second pass: match on display name stem
        stem = pst_path.stem.lower()
        for i in range(1, stores.Count + 1):  # type: ignore[attr-defined]
            store = stores.Item(i)
            try:
                if store.DisplayName.lower() == stem:
                    return store.DisplayName
            except Exception:
                continue
        return None

    # --- Enumeration ----------------------------------------------------------------
    def iter_mail_items(self, store_display_name: str, include_non_mail: bool = False, cancel_event: Optional[Event] = None) -> Iterator[MailItemInfo]:
        """Yield MailItemInfo for all items in the given store.

        This is a simplified example: real implementation must traverse folders.
        """
        # Import COM optimization modules
        import pythoncom
        import gc
        
        # Initialize COM for this thread if needed
        try:
            pythoncom.CoInitialize()
        except:
            pass  # Already initialized
        
        stores = self._namespace.Stores
        for i in range(1, stores.Count + 1):  # type: ignore[attr-defined]
            store = stores.Item(i)
            if store.DisplayName == store_display_name:
                logging.info("Enumerating items in store: %s", store.DisplayName)
                # Pre-enumeration can be skipped (fast mode) for performance
                fast_mode = getattr(self, 'fast_enumeration', False)
                turbo_mode = getattr(self, 'turbo_mode', False)
                
                if not fast_mode and not turbo_mode:
                    pre_stack: List[tuple[object, str]] = [(store.GetRootFolder(), "")]
                    folder_paths: List[str] = []
                    while pre_stack:
                        # Check for cancellation during pre-enumeration
                        if cancel_event and cancel_event.is_set():
                            logging.info("Enumeration cancelled during pre-enumeration")
                            return
                        f0, rel0 = pre_stack.pop()
                        folder_paths.append(rel0 or "/")
                        try:
                            sub_cnt0 = f0.Folders.Count  # type: ignore[attr-defined]
                        except Exception:
                            sub_cnt0 = 0
                        for k in range(1, sub_cnt0 + 1):  # type: ignore[attr-defined]
                            sub0 = f0.Folders.Item(k)  # type: ignore
                            rel_child = f"{rel0}/{sub0.Name}" if rel0 else sub0.Name
                            pre_stack.append((sub0, rel_child))
                    logging.debug("Pre-enumeration: %s folders discovered", len(folder_paths))
                    logging.debug("First folders: %s", folder_paths[:10])  # Feature E
                root = store.GetRootFolder()
                stack: List[tuple[object, str]] = [(root, "")]  # (folder, relative_path)
                folder_counter = 0
                item_counter = 0
                skipped_non_mail = 0
                last_speed_log = 0
                start_time = datetime.now()
                
                while stack:
                    # Check for cancellation at start of each folder
                    if cancel_event and cancel_event.is_set():
                        logging.info("Enumeration cancelled after processing %s folders, %s items", folder_counter, item_counter)
                        return
                    f, rel_path = stack.pop()
                    folder_counter += 1
                    
                    # Log folder processing with timing info
                    if folder_counter % 10 == 0:
                        elapsed = (datetime.now() - start_time).total_seconds()
                        rate = item_counter / elapsed if elapsed > 0 else 0
                        logging.debug("Processing folder %s (%s): %s items/sec", folder_counter, rel_path or "/", f"{rate:.1f}")
                    
                    try:
                        sub_count = f.Folders.Count  # type: ignore[attr-defined]
                    except Exception:
                        sub_count = 0
                    for j in range(1, sub_count + 1):  # type: ignore[attr-defined]
                        sub = f.Folders.Item(j)  # type: ignore
                        sub_rel = f"{rel_path}/{sub.Name}" if rel_path else sub.Name
                        stack.append((sub, sub_rel))
                    
                    try:
                        item_count = f.Items.Count  # type: ignore[attr-defined]
                    except Exception:
                        item_count = 0
                    
                    # Skip empty folders
                    if item_count == 0:
                        continue
                    
                    # Optimize for large folders with aggressive batching and COM cleanup
                    if item_count == 0:
                        batch_size = 1  # Prevent zero batch size
                    elif turbo_mode:
                        batch_size = min(2000, item_count)  # Extra large batches in turbo mode
                    elif item_count > 10000:
                        batch_size = 1000  # Much larger batches for very large folders
                    elif item_count > 5000:
                        batch_size = 500
                    else:
                        batch_size = min(200, item_count)
                    
                    # Import gc once for large folders
                    if item_count > 5000:
                        import gc
                    
                    for batch_start in range(1, item_count + 1, batch_size):
                        batch_end = min(batch_start + batch_size - 1, item_count)
                        
                        # Check for cancellation at start of each batch
                        if cancel_event and cancel_event.is_set():
                            logging.info("Enumeration cancelled after processing %s folders, %s items", folder_counter, item_counter)
                            return
                        
                        # For very large folders, use minimal property access mode
                        minimal_mode = turbo_mode or item_count > 5000  # Turbo mode always uses minimal properties
                        
                        # Process batch of items with optimized COM handling
                        batch_processed = 0
                        
                        # For large batches, pre-collect items to reduce COM overhead
                        if batch_size > 200:
                            # Collect batch items first
                            batch_items = []
                            try:
                                items_collection = f.Items  # Cache the Items collection  # type: ignore[attr-defined]
                                for j in range(batch_start, batch_end + 1):
                                    try:
                                        item = items_collection.Item(j)  # type: ignore[attr-defined]
                                        batch_items.append(item)
                                    except:
                                        continue
                            except:
                                # Fall back to individual access
                                batch_items = []
                                for j in range(batch_start, batch_end + 1):
                                    try:
                                        batch_items.append(f.Items.Item(j))  # type: ignore[attr-defined]
                                    except:
                                        continue
                        else:
                            batch_items = []
                        # Process items in batch
                        process_range = batch_items if batch_items else range(batch_start, batch_end + 1)
                        
                        for item_ref in process_range:
                            # Check for cancellation every 10 items for responsive cancellation
                            if batch_processed % 10 == 0 and cancel_event and cancel_event.is_set():
                                logging.info("Enumeration cancelled during batch processing after %s items", item_counter)
                                return
                                
                            try:
                                # Get the item - either from pre-collected batch or by index
                                if batch_items:
                                    it = item_ref
                                else:
                                    it = f.Items.Item(item_ref)  # type: ignore
                                
                                # Fast property access with minimal error handling for speed
                                try:
                                    # In turbo mode, access absolute minimum properties for maximum speed
                                    if turbo_mode:
                                        # Ultra-minimal property set - only what's absolutely required
                                        entry_id = str(getattr(it, "EntryID", "") or "")
                                        size = int(getattr(it, "Size", 0) or 0)
                                        # Skip all optional properties in turbo mode
                                        message_class = ""
                                        subject = ""
                                        sender_email = ""
                                        received = None
                                    elif minimal_mode:
                                        # Minimal property set for maximum speed
                                        entry_id = str(it.EntryID or "")  # type: ignore[attr-defined]
                                        size = int(getattr(it, "Size", 0) or 0)
                                        message_class = str(getattr(it, "MessageClass", "") or "")
                                        subject = ""  # Skip subject for speed
                                        sender_email = ""  # Skip sender for speed
                                        received = None  # Skip received for speed
                                    else:
                                        # Access all properties at once to minimize COM calls
                                        # Use getattr with defaults to handle missing properties gracefully
                                        entry_id = str(getattr(it, "EntryID", "") or "")
                                        size = int(getattr(it, "Size", 0) or 0)
                                        
                                        # Safe datetime handling to avoid win32timezone issues
                                        try:
                                            received = getattr(it, "ReceivedTime", None)
                                            # Convert to naive datetime if needed
                                            if received and hasattr(received, 'replace'):
                                                try:
                                                    received = received.replace(tzinfo=None)
                                                except:
                                                    received = None
                                        except Exception:
                                            received = None  # Fallback if timezone conversion fails
                                            
                                        message_class = str(getattr(it, "MessageClass", "") or "")
                                        subject = str(getattr(it, "Subject", "") or "")
                                        sender_email = str(getattr(it, "SenderEmailAddress", "") or "")
                                    
                                    # Explicitly release COM object reference
                                    it = None
                                    
                                except Exception as e:
                                    # Handle specific COM errors more gracefully
                                    error_code = getattr(e, 'winerror', None) if hasattr(e, 'winerror') else str(e)
                                    
                                    # Don't log common/expected errors unless in debug mode
                                    if error_code == -2147467262:  # E_NOINTERFACE - corrupted/unsupported item
                                        if not turbo_mode:  # Only log in non-turbo mode for less noise
                                            logging.debug("Skipping corrupted item (E_NOINTERFACE)")
                                    elif error_code == -2147023174:  # Access denied
                                        logging.debug("Skipping protected item (access denied)")
                                    else:
                                        # Log unexpected errors for debugging
                                        logging.debug("Error accessing item properties: %s", e)
                                    
                                    it = None
                                    continue
                                
                                # Feature D: filter for primary mail class IPM.Note (still include if empty)
                                # Skip filtering in turbo mode for maximum speed
                                if (not turbo_mode) and (not include_non_mail) and message_class and not message_class.startswith("IPM.Note"):
                                    skipped_non_mail += 1
                                    continue
                                
                                if received:
                                    try:
                                        received_py = datetime.fromtimestamp(received.timestamp())  # type: ignore[attr-defined]
                                    except Exception:  # pragma: no cover
                                        received_py = None
                                else:
                                    received_py = None
                                
                                yield MailItemInfo(
                                    entry_id=entry_id,
                                    subject=subject,
                                    size=size,
                                    received=received_py,
                                    folder_path=rel_path or "",
                                    sender_email=sender_email,
                                )
                                item_counter += 1
                                batch_processed += 1
                                
                            except Exception as e:  # pragma: no cover
                                logging.debug("Skipping item due to error: %s", e)
                        
                        # Aggressive memory management for large batches
                        if item_count > 5000 and batch_processed > 0:
                            gc.collect()  # Force garbage collection to free COM objects
                            
                            # For very large folders, also clear batch items to free memory
                            if batch_items:
                                for item in batch_items:
                                    try:
                                        item = None
                                    except:
                                        pass
                                batch_items.clear()
                        
                        # Log progress with timing info every 500 items for large folders
                        progress_interval = 500 if item_count > 5000 else 1000
                        if item_counter > 0 and item_counter - last_speed_log >= progress_interval:
                            elapsed = (datetime.now() - start_time).total_seconds()
                            rate = item_counter / elapsed if elapsed > 0 else 0
                            logging.info("Processed %s items in %s folders (%.1f items/sec)", 
                                       item_counter, folder_counter, rate)
                            last_speed_log = item_counter
                logging.info("Enumeration complete: %s folders, %s items", folder_counter, item_counter)
                if skipped_non_mail:
                    logging.info("Skipped non-mail items (message class != IPM.Note*): %s", skipped_non_mail)
                
                # Final COM cleanup
                try:
                    gc.collect()  # Force garbage collection of COM objects
                    pythoncom.CoUninitialize()
                except:
                    pass
                return
        raise RuntimeError(f"Store not found: {store_display_name}")

    # --- Copy / Move -----------------------------------------------------------------
    def _ensure_folder_path(self, store, path: str):  # pragma: no cover - Outlook specific
        if not path:
            return store.GetRootFolder()
        parts = path.split('/')
        current = store.GetRootFolder()
        store_name = store.DisplayName
        for part in parts:
            # Build progressive path key
            prefix_index = parts.index(part) + 1
            progressive = '/'.join(parts[:prefix_index])
            cache_key = (store_name, progressive)
            if cache_key in self._folder_cache:
                current = self._folder_cache[cache_key]
                continue
            found = None
            for i in range(1, current.Folders.Count + 1):  # type: ignore[attr-defined]
                f = current.Folders.Item(i)  # type: ignore[attr-defined]
                if f.Name == part:
                    found = f
                    break
            if not found:
                current.Folders.Add(part)  # type: ignore[attr-defined]
                # Re-search to get the created folder object
                for i in range(1, current.Folders.Count + 1):  # type: ignore[attr-defined]
                    f = current.Folders.Item(i)  # type: ignore[attr-defined]
                    if f.Name == part:
                        found = f
                        break
            current = found or current
            self._folder_cache[cache_key] = current
        return current

    def copy_item_to_store(self, entry_id: str, target_store_name: str, folder_path: str, target_pst_path: Optional[Path] = None, move: bool = False) -> None:
        """Copy item preserving (creating) folder structure with retry logic.

        Retries locating the destination store (display name or file path) with
        exponential backoff to handle Outlook delays after AddStore.
        """
        try:
            item = self._namespace.GetItemFromID(entry_id)
        except Exception as e:  # pragma: no cover
            logging.debug("Failed to resolve source item %s: %s", entry_id, e)
            return
        attempts = 10
        dest_store = None
        for attempt in range(1, attempts + 1):
            try:
                path_match = None
                name_match = None
                for store in self._namespace.Stores:  # type: ignore
                    if target_pst_path is not None:
                        try:
                            if Path(store.FilePath).resolve() == target_pst_path.resolve():
                                path_match = store
                                break  # exact path match wins immediately
                        except Exception:
                            pass
                    if store.DisplayName == target_store_name and name_match is None:
                        name_match = store  # keep first name match as fallback
                dest_store = path_match or name_match
                if dest_store:
                    break
            except Exception:
                pass
            if attempt < attempts:
                delay = 0.2 * attempt
                logging.debug("Dest store %s not ready (attempt %s/%s), sleeping %.1fs", target_store_name, attempt, attempts, delay)
                import time as _t
                _t.sleep(delay)
        if not dest_store:
            logging.warning("Target store %s not found for copy after retries", target_store_name)
            return
        try:
            target_folder = self._ensure_folder_path(dest_store, folder_path)
            if move:
                # Move original item
                item.Move(target_folder)  # type: ignore[attr-defined]
            else:
                new_item = item.Copy()  # type: ignore[attr-defined]
                new_item.Move(target_folder)  # type: ignore[attr-defined]
            try:
                pst_path_used = Path(getattr(dest_store, 'FilePath', ''))
            except Exception:
                pst_path_used = None
            if pst_path_used and target_pst_path and pst_path_used.resolve() != target_pst_path.resolve():
                logging.debug(
                    "Copied item %s to store display '%s' but path mismatch (expected %s got %s)",
                    entry_id,
                    target_store_name,
                    target_pst_path,
                    pst_path_used,
                )
            else:
                if not self.suppress_item_logs:
                    logging.debug("Copied item %s to %s (%s)/%s", entry_id, target_store_name, target_pst_path, folder_path)
        except Exception as e:  # pragma: no cover
            error_msg = str(e)
            error_code = getattr(e, 'winerror', None) if hasattr(e, 'winerror') else None
            
            # Check for PST size limit error
            if (error_code == -2147219956 or 
                "maximum size" in error_msg.lower() or 
                "reached the maximum size" in error_msg.lower()):
                logging.warning("PST size limit reached during copy of item %s. Attempting recovery...", entry_id)
                raise PSTSizeExceededException(f"Source PST has reached maximum size while copying item {entry_id}: {e}")
            else:
                logging.debug("Failed to copy item %s: %s", entry_id, e)
                raise

    # --- Cleanup ---------------------------------------------------------------------
    def close(self) -> None:
        logging.info("Closing Outlook session")
        # Rely on COM object release by refcount.


def is_outlook_available() -> bool:
    """Return True if Outlook COM dispatch is available."""
    return Dispatch is not None


__all__ = ["OutlookSession", "MailItemInfo", "is_outlook_available"]
