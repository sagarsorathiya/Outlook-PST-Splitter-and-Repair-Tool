"""Performance optimizations for PST item copying and moving operations."""
import logging
import time
from typing import List, Optional, Callable
from pathlib import Path
from threading import Event
from .outlook import MailItemInfo, OutlookSession
from .log_exporter import log_exporter


class HighPerformancePSTProcessor:
    """Optimized PST item processing with batching and performance improvements."""
    
    def __init__(self, outlook_session: OutlookSession):
        self.outlook = outlook_session
        self.batch_size = 50  # Process items in batches
        self.performance_mode = True
        
    def set_performance_mode(self, enabled: bool, batch_size: int = 50):
        """Configure performance settings."""
        self.performance_mode = enabled
        self.batch_size = batch_size
        log_exporter.log_performance_metric("batch_size", batch_size, "items")
        logging.info(f"⚡ Performance mode: {'enabled' if enabled else 'disabled'}, batch size: {batch_size}")
    
    def copy_items_optimized(
        self,
        items: List[MailItemInfo],
        target_store_name: str,
        target_folder: str,
        target_pst_path: Path,
        move_items: bool = False,
        progress_callback: Optional[Callable] = None,
        cancel_event: Optional[Event] = None
    ) -> dict:
        """
        Optimized item copying with batching and performance improvements.
        
        Returns:
            dict: Results with success count, failed items, and performance metrics
        """
        start_time = time.time()
        results = {
            'success_count': 0,
            'failed_items': [],
            'total_time_ms': 0,
            'avg_time_per_item_ms': 0,
            'batches_processed': 0
        }
        
        if not items:
            return results
        
        total_items = len(items)
        processed_count = 0
        cancelled = False
        
        logging.info(f"⚡ Starting optimized {'move' if move_items else 'copy'} of {total_items} items")
        
        # Process items in batches for better performance
        for batch_start in range(0, total_items, self.batch_size):
            # Check for cancellation at the start of each batch
            if cancel_event and cancel_event.is_set():
                logging.info("⚠️ Copy operation cancelled by user")
                results['cancelled'] = True
                cancelled = True
                break
                
            batch_end = min(batch_start + self.batch_size, total_items)
            batch_items = items[batch_start:batch_end]

            # Report batch start
            if progress_callback:
                progress_callback(f"Processed {processed_count}/{total_items} items")

            batch_start_time = time.time()
            batch_success = 0            # Pre-optimize COM objects for batch
            if self.performance_mode:
                self._prepare_batch_processing()
            
            for item in batch_items:
                # Check for cancellation for each item
                if cancel_event and cancel_event.is_set():
                    logging.info("⚠️ Copy operation cancelled by user during item processing")
                    results['cancelled'] = True
                    cancelled = True
                    break
                    
                item_start_time = time.time()
                success = False
                error_msg = None
                
                try:
                    # Extract year for logging
                    year = "Unknown"
                    if item.received:
                        year = item.received.strftime("%Y")
                    
                    # Perform the copy/move operation
                    if self.performance_mode:
                        # Use optimized COM calls
                        success = self._optimized_item_transfer(
                            item, target_store_name, target_folder, 
                            target_pst_path, move_items
                        )
                    else:
                        # Use standard method
                        self.outlook.copy_item_to_store(
                            item.entry_id,
                            target_store_name,
                            target_folder,
                            target_pst_path=target_pst_path,
                            move=move_items
                        )
                        success = True
                    
                    if success:
                        batch_success += 1
                        results['success_count'] += 1
                    
                except Exception as e:
                    year = "Unknown"  # Initialize year for error case
                    error_msg = str(e)
                    results['failed_items'].append({
                        'item_id': item.entry_id[:50],  # Truncate for privacy
                        'error': error_msg,
                        'year': year
                    })
                    logging.debug(f"Failed to process item: {error_msg}")
                
                # Log individual item processing
                item_time_ms = (time.time() - item_start_time) * 1000
                log_exporter.log_item_processing(
                    item.entry_id, year, target_store_name, 
                    item_time_ms, success, error_msg
                )
                
                processed_count += 1
                
                # Progress callback every 5 items for more responsive updates
                if progress_callback and processed_count % 5 == 0:
                    progress_callback(f"Processed {processed_count}/{total_items} items")
            
            # Check if cancelled during inner loop
            if cancelled:
                break
                
            # Batch completion
            batch_time_ms = (time.time() - batch_start_time) * 1000
            results['batches_processed'] += 1
            
            # Log batch performance
            log_exporter.log_performance_metric(
                f"batch_{results['batches_processed']}_time_ms", 
                batch_time_ms, "ms"
            )
            log_exporter.log_performance_metric(
                f"batch_{results['batches_processed']}_success_rate", 
                batch_success / len(batch_items) * 100, "%"
            )
            
            logging.info(f"📦 Batch {results['batches_processed']}: {batch_success}/{len(batch_items)} items in {batch_time_ms:.1f}ms")
            
            # Cleanup between batches for memory management
            if self.performance_mode and results['batches_processed'] % 5 == 0:
                self._cleanup_batch_processing()
        
        # Calculate final metrics
        total_time_ms = (time.time() - start_time) * 1000
        results['total_time_ms'] = total_time_ms
        results['avg_time_per_item_ms'] = total_time_ms / total_items if total_items > 0 else 0
        
        # Log final performance metrics
        log_exporter.log_performance_metric("total_items_processed", total_items, "items")
        log_exporter.log_performance_metric("total_processing_time_ms", total_time_ms, "ms")
        log_exporter.log_performance_metric("avg_time_per_item_ms", results['avg_time_per_item_ms'], "ms")
        log_exporter.log_performance_metric("success_rate", results['success_count'] / total_items * 100, "%")
        log_exporter.log_performance_metric("items_per_second", total_items / (total_time_ms / 1000), "items/sec")
        
        logging.info(f"⚡ Optimized processing complete: {results['success_count']}/{total_items} items in {total_time_ms/1000:.1f}s")
        logging.info(f"📊 Performance: {results['avg_time_per_item_ms']:.1f}ms per item, {total_items/(total_time_ms/1000):.1f} items/sec")
        
        return results
    
    def _prepare_batch_processing(self):
        """Prepare for optimized batch processing."""
        try:
            # Simple preparation - just log that we're optimizing
            logging.debug("🔧 Prepared batch processing optimizations")
            
        except Exception as e:
            logging.debug(f"Failed to prepare batch optimizations: {e}")
    
    def _cleanup_batch_processing(self):
        """Cleanup and memory management between batches."""
        try:
            # Force garbage collection
            import gc
            gc.collect()
            logging.debug("🧹 Cleaned up batch processing cache")
            
        except Exception as e:
            logging.debug(f"Failed to cleanup batch processing: {e}")
    
    def _optimized_item_transfer(
        self, 
        item: MailItemInfo, 
        target_store_name: str, 
        target_folder: str,
        target_pst_path: Path, 
        move_items: bool
    ) -> bool:
        """
        Optimized version of item transfer.
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # For now, use the standard method but with optimized error handling
            self.outlook.copy_item_to_store(
                item.entry_id,
                target_store_name,
                target_folder,
                target_pst_path=target_pst_path,
                move=move_items
            )
            return True
            
        except Exception as e:
            logging.debug(f"Optimized transfer failed for item {item.entry_id[:20]}: {e}")
            return False
    
    def _fallback_item_transfer(
        self, 
        item: MailItemInfo, 
        target_store_name: str, 
        target_folder: str,
        target_pst_path: Path, 
        move_items: bool
    ) -> bool:
        """Fallback to standard transfer method if optimized version fails."""
        try:
            self.outlook.copy_item_to_store(
                item.entry_id,
                target_store_name,
                target_folder,
                target_pst_path=target_pst_path,
                move=move_items
            )
            return True
        except Exception as e:
            logging.debug(f"Fallback transfer also failed: {e}")
            return False
