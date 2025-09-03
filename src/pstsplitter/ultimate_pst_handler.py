"""
Ultimate PST Handler - Practical solutions for critically full PSTs
Uses alternative approaches that don't require source PST space
"""
import os
import sqlite3
import tempfile
from pathlib import Path
from typing import Dict, List, Any, Optional
import logging
import json
from datetime import datetime
import shutil

logger = logging.getLogger(__name__)


class UltimatePSTHandler:
    """
    Handles critically full PSTs using practical alternatives that don't require source space
    """
    
    def __init__(self):
        self.temp_dir = tempfile.mkdtemp(prefix="pst_ultimate_")
        self.metadata_db = os.path.join(self.temp_dir, "pst_metadata.db")
        self.init_metadata_db()
        self.consecutive_failures = 0
        self.max_consecutive_failures = 5
    
    def init_metadata_db(self):
        """Initialize SQLite database for item metadata tracking"""
        try:
            conn = sqlite3.connect(self.metadata_db)
            cursor = conn.cursor()
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS mail_items (
                    id INTEGER PRIMARY KEY,
                    entry_id TEXT UNIQUE,
                    subject TEXT,
                    sender TEXT,
                    received_time TEXT,
                    size_bytes INTEGER,
                    year INTEGER,
                    month INTEGER,
                    folder_path TEXT,
                    processed BOOLEAN DEFAULT FALSE
                )
            ''')
            
            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_year ON mail_items(year);
                CREATE INDEX IF NOT EXISTS idx_entry_id ON mail_items(entry_id);
                CREATE INDEX IF NOT EXISTS idx_processed ON mail_items(processed);
            ''')
            
            conn.commit()
            conn.close()
        except Exception as e:
            logger.warning(f"Failed to initialize metadata DB: {e}")
    
    def handle_critically_full_pst(
        self, 
        source_path: str, 
        output_dir: str, 
        split_by: str = "year",
        progress_callback=None
    ) -> Dict[str, Any]:
        """
        Handle critically full PST using practical methods that don't require source space
        """
        results = {
            'success': False,
            'method_used': None,
            'files_created': [],
            'items_processed': 0,
            'errors': []
        }
        
        logger.info("🚨 ULTIMATE PST HANDLER ACTIVATED for critically full PST")
        
        # Strategy 1: Move-Only Mode (No Copy Operations)
        try:
            logger.info("🔄 Attempting Move-Only strategy (no copy operations)")
            if progress_callback:
                progress_callback("Using Move-Only strategy...")
                
            move_result = self._move_only_strategy(source_path, output_dir, split_by, progress_callback)
            if move_result['success']:
                results.update(move_result)
                results['method_used'] = 'move_only'
                logger.info("✅ Move-Only strategy succeeded!")
                return results
        except Exception as e:
            logger.warning(f"Move-Only strategy failed: {e}")
            results['errors'].append(f"Move-Only: {e}")
        
        # Strategy 2: Export to MSG files then Reimport
        try:
            logger.info("📤 Attempting Export-Reimport strategy")
            if progress_callback:
                progress_callback("Using Export-Reimport strategy...")
                
            export_result = self._export_reimport_strategy(source_path, output_dir, split_by, progress_callback)
            if export_result['success']:
                results.update(export_result)
                results['method_used'] = 'export_reimport'
                logger.info("✅ Export-Reimport strategy succeeded!")
                return results
        except Exception as e:
            logger.warning(f"Export-Reimport strategy failed: {e}")
            results['errors'].append(f"Export-Reimport: {e}")
        
        # Strategy 3: Direct MAPI Access (Bypass normal operations)
        try:
            logger.info("⚡ Attempting Direct MAPI access")
            if progress_callback:
                progress_callback("Using Direct MAPI access...")
                
            mapi_result = self._direct_mapi_strategy(source_path, output_dir, split_by, progress_callback)
            if mapi_result['success']:
                results.update(mapi_result)
                results['method_used'] = 'direct_mapi'
                logger.info("✅ Direct MAPI strategy succeeded!")
                return results
        except Exception as e:
            logger.warning(f"Direct MAPI strategy failed: {e}")
            results['errors'].append(f"Direct MAPI: {e}")
        
        # Strategy 4: Item-by-Item with Skip on Error
        try:
            logger.info("🎯 Attempting Item-by-Item with skip strategy")
            if progress_callback:
                progress_callback("Using Item-by-Item skip strategy...")
                
            skip_result = self._item_by_item_skip_strategy(source_path, output_dir, split_by, progress_callback)
            if skip_result['success']:
                results.update(skip_result)
                results['method_used'] = 'item_by_item_skip'
                logger.info("✅ Item-by-Item skip strategy succeeded!")
                return results
        except Exception as e:
            logger.warning(f"Item-by-Item skip strategy failed: {e}")
            results['errors'].append(f"Item-by-Item: {e}")
        
        logger.error("❌ All ultimate strategies failed")
        return results
    
    def _move_only_strategy(self, source_path: str, output_dir: str, split_by: str, progress_callback) -> Dict[str, Any]:
        """
        Move items directly without copying - completely bypasses space issues
        """
        from .outlook import OutlookSession
        
        results = {'success': False, 'files_created': [], 'items_processed': 0}
        outlook = None
        
        try:
            outlook = OutlookSession()
            
            # Attach source PST
            outlook.attach_pst(Path(source_path))
            source_store_name = outlook.find_store_by_path(Path(source_path))
            
            if not source_store_name:
                raise Exception("Failed to find source store")
            
            # Create target PSTs for each period
            target_stores = {}
            
            # Get items using enumeration
            items = list(outlook.iter_mail_items(source_store_name, include_non_mail=False))
            logger.info(f"Found {len(items)} items to move")
            
            for idx, item in enumerate(items):
                try:
                    # Get year for grouping
                    try:
                        received_time = item.received
                        if received_time:
                            year = received_time.year
                        else:
                            year = 'unknown'
                    except:
                        year = 'unknown'
                    
                    # Create target PST if needed
                    if year not in target_stores:
                        target_filename = f"{Path(source_path).stem}_{year}.pst"
                        target_path = os.path.join(output_dir, target_filename)
                        
                        if not os.path.exists(target_path):
                            outlook.create_new_pst(Path(target_path))
                            results['files_created'].append(target_path)
                        
                        target_stores[year] = target_path
                    
                    # Move item using move operation instead of copy
                    try:
                        target_store_name = outlook.find_store_by_path(Path(target_stores[year]))
                        if target_store_name:
                            outlook.copy_item_to_store(
                                item.entry_id,
                                target_store_name,
                                "Items",
                                target_pst_path=Path(target_stores[year]),
                                move=True  # This is the key - MOVE instead of COPY
                            )
                            results['items_processed'] += 1
                    except Exception as move_error:
                        logger.debug(f"Failed to move item {idx}: {move_error}")
                        continue
                    
                    if idx % 50 == 0 and progress_callback:
                        progress_callback(f"Moved {results['items_processed']} of {len(items)} items")
                        
                except Exception as e:
                    logger.debug(f"Failed to process item {idx}: {e}")
                    continue
            
            results['success'] = results['items_processed'] > 0
            logger.info(f"Move-Only strategy processed {results['items_processed']} items")
            
        except Exception as e:
            logger.error(f"Move-Only strategy failed: {e}")
            raise
        finally:
            if outlook:
                try:
                    outlook.detach_pst(Path(source_path))
                    for target_path in results['files_created']:
                        try:
                            outlook.detach_pst(Path(target_path))
                        except:
                            pass
                except:
                    pass
        
        return results
    
    def _export_reimport_strategy(self, source_path: str, output_dir: str, split_by: str, progress_callback) -> Dict[str, Any]:
        """
        Export to MSG files, then reimport to new PSTs - completely bypasses space issues
        """
        results = {'success': False, 'files_created': [], 'items_processed': 0}
        outlook = None
        export_dir = None
        
        try:
            import win32com.client
            
            # Create Outlook application
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            # Add source PST
            namespace.AddStore(source_path)
            source_store = None
            
            for store in namespace.Stores:
                if store.FilePath and store.FilePath.lower() == source_path.lower():
                    source_store = store
                    break
            
            if not source_store:
                raise Exception("Could not find source store")
            
            # Create temporary export directory
            export_dir = os.path.join(self.temp_dir, "exported_msgs")
            os.makedirs(export_dir, exist_ok=True)
            
            exported_by_year = {}
            
            def export_folder_items(folder, folder_path=""):
                """Recursively export items from folder"""
                try:
                    items = folder.Items
                    for i in range(1, items.Count + 1):
                        try:
                            item = items[i]
                            
                            # Get year
                            try:
                                received_time = getattr(item, 'ReceivedTime', None)
                                year = received_time.year if received_time else 'unknown'
                            except:
                                year = 'unknown'
                            
                            # Create year directory
                            year_dir = os.path.join(export_dir, str(year))
                            os.makedirs(year_dir, exist_ok=True)
                            
                            # Export to MSG file
                            msg_filename = f"item_{results['items_processed']:06d}.msg"
                            msg_path = os.path.join(year_dir, msg_filename)
                            
                            # Save as MSG file
                            item.SaveAs(msg_path, 3)  # 3 = olMSG format
                            
                            if year not in exported_by_year:
                                exported_by_year[year] = []
                            exported_by_year[year].append(msg_path)
                            
                            results['items_processed'] += 1
                            
                            if results['items_processed'] % 25 == 0 and progress_callback:
                                progress_callback(f"Exported {results['items_processed']} items")
                                
                        except Exception as e:
                            logger.debug(f"Failed to export item: {e}")
                            continue
                    
                    # Process subfolders
                    for subfolder in folder.Folders:
                        export_folder_items(subfolder, f"{folder_path}/{subfolder.Name}")
                        
                except Exception as e:
                    logger.debug(f"Failed to process folder {folder_path}: {e}")
            
            # Export all items
            root_folder = source_store.GetRootFolder()
            export_folder_items(root_folder)
            
            # Create new PSTs and import MSG files
            for year, msg_files in exported_by_year.items():
                if not msg_files:
                    continue
                    
                if progress_callback:
                    progress_callback(f"Creating PST for year {year} ({len(msg_files)} items)")
                
                target_filename = f"{Path(source_path).stem}_{year}.pst"
                target_path = os.path.join(output_dir, target_filename)
                
                # Create new PST
                namespace.AddStore(target_path)
                target_store = None
                
                for store in namespace.Stores:
                    if store.FilePath and store.FilePath.lower() == target_path.lower():
                        target_store = store
                        break
                
                if target_store:
                    target_folder = target_store.GetDefaultFolder(6)  # olFolderInbox
                    
                    for msg_file in msg_files:
                        try:
                            # Import MSG file
                            imported_item = namespace.OpenSharedItem(msg_file)
                            if imported_item:
                                imported_item.Move(target_folder)
                        except Exception as e:
                            logger.debug(f"Failed to import {msg_file}: {e}")
                    
                    results['files_created'].append(target_path)
            
            results['success'] = len(results['files_created']) > 0
            logger.info(f"Export-Reimport strategy created {len(results['files_created'])} PST files")
            
        except Exception as e:
            logger.error(f"Export-Reimport strategy failed: {e}")
            raise
        finally:
            if outlook:
                try:
                    outlook.Quit()
                except:
                    pass
            # Clean up temporary MSG files
            if export_dir:
                try:
                    shutil.rmtree(export_dir)
                except:
                    pass
        
        return results
    
    def _direct_mapi_strategy(self, source_path: str, output_dir: str, split_by: str, progress_callback) -> Dict[str, Any]:
        """
        Use direct MAPI access to bypass normal Outlook operations
        """
        results = {'success': False, 'files_created': [], 'items_processed': 0}
        outlook = None
        
        try:
            import win32com.client
            from win32com.client import gencache
            
            # Create Outlook application with different approach
            outlook = gencache.EnsureDispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            # Add source PST
            namespace.AddStore(source_path)
            source_store = None
            
            for store in namespace.Stores:
                if store.FilePath and store.FilePath.lower() == source_path.lower():
                    source_store = store
                    break
            
            if not source_store:
                raise Exception("Could not find source store")
            
            # Collect items by year using MAPI
            items_by_year = {}
            
            def collect_items_from_folder(folder):
                """Collect items using MAPI entry IDs"""
                try:
                    items = folder.Items
                    for i in range(1, items.Count + 1):
                        try:
                            item = items[i]
                            
                            # Get EntryID for MAPI access
                            entry_id = item.EntryID
                            store_id = source_store.StoreID
                            
                            # Get year
                            try:
                                received_time = getattr(item, 'ReceivedTime', None)
                                year = received_time.year if received_time else 'unknown'
                            except:
                                year = 'unknown'
                            
                            if year not in items_by_year:
                                items_by_year[year] = []
                            
                            items_by_year[year].append({
                                'entry_id': entry_id,
                                'store_id': store_id
                            })
                            
                            results['items_processed'] += 1
                            
                            if results['items_processed'] % 100 == 0 and progress_callback:
                                progress_callback(f"Collected {results['items_processed']} items")
                                
                        except Exception as e:
                            logger.debug(f"Failed to collect item: {e}")
                            continue
                    
                    # Process subfolders
                    for subfolder in folder.Folders:
                        collect_items_from_folder(subfolder)
                        
                except Exception as e:
                    logger.debug(f"Failed to collect from folder: {e}")
            
            # Collect all items
            root_folder = source_store.GetRootFolder()
            collect_items_from_folder(root_folder)
            
            # Create target PSTs and move items using MAPI
            for year, items in items_by_year.items():
                if not items:
                    continue
                    
                if progress_callback:
                    progress_callback(f"Creating PST for year {year} ({len(items)} items)")
                
                target_filename = f"{Path(source_path).stem}_{year}.pst"
                target_path = os.path.join(output_dir, target_filename)
                
                # Create new PST
                namespace.AddStore(target_path)
                target_store = None
                
                for store in namespace.Stores:
                    if store.FilePath and store.FilePath.lower() == target_path.lower():
                        target_store = store
                        break
                
                if target_store:
                    target_folder = target_store.GetDefaultFolder(6)  # olFolderInbox
                    
                    for item_info in items:
                        try:
                            # Get item using MAPI
                            mapi_item = namespace.GetItemFromID(
                                item_info['entry_id'], 
                                item_info['store_id']
                            )
                            
                            # Move to target folder
                            if mapi_item:
                                mapi_item.Move(target_folder)
                                
                        except Exception as e:
                            logger.debug(f"Failed to move item via MAPI: {e}")
                            continue
                    
                    results['files_created'].append(target_path)
            
            results['success'] = len(results['files_created']) > 0
            logger.info(f"Direct MAPI strategy created {len(results['files_created'])} PST files")
            
        except Exception as e:
            logger.error(f"Direct MAPI strategy failed: {e}")
            raise
        finally:
            if outlook:
                try:
                    outlook.Quit()
                except:
                    pass
        
        return results
    
    def _item_by_item_skip_strategy(self, source_path: str, output_dir: str, split_by: str, progress_callback) -> Dict[str, Any]:
        """
        Process items one by one, skipping problematic items to make progress
        """
        from .outlook import OutlookSession
        
        results = {'success': False, 'files_created': [], 'items_processed': 0}
        skipped_items = 0
        outlook = None
        
        try:
            outlook = OutlookSession()
            
            # Attach source PST
            outlook.attach_pst(Path(source_path))
            source_store_name = outlook.find_store_by_path(Path(source_path))
            
            if not source_store_name:
                raise Exception("Failed to find source store")
            
            # Create target PSTs for each period
            target_stores = {}
            
            # Get items using enumeration
            items = list(outlook.iter_mail_items(source_store_name, include_non_mail=False))
            logger.info(f"Processing {len(items)} items with skip strategy")
            
            for idx, item in enumerate(items):
                try:
                    # Get year for grouping
                    try:
                        received_time = item.received
                        if received_time:
                            year = received_time.year
                        else:
                            year = 'unknown'
                    except:
                        year = 'unknown'
                    
                    # Create target PST if needed
                    if year not in target_stores:
                        target_filename = f"{Path(source_path).stem}_{year}.pst"
                        target_path = os.path.join(output_dir, target_filename)
                        
                        if not os.path.exists(target_path):
                            outlook.create_new_pst(Path(target_path))
                            results['files_created'].append(target_path)
                        
                        target_stores[year] = target_path
                    
                    # Try to copy item, skip on any error
                    try:
                        target_store_name = outlook.find_store_by_path(Path(target_stores[year]))
                        if target_store_name:
                            outlook.copy_item_to_store(
                                item.entry_id,
                                target_store_name,
                                "Items",
                                target_pst_path=Path(target_stores[year]),
                                move=False  # Copy first, then can try move if needed
                            )
                            results['items_processed'] += 1
                    except Exception as copy_error:
                        # Skip this item and continue - don't let one item stop the whole process
                        skipped_items += 1
                        logger.debug(f"Skipped problematic item {idx}: {copy_error}")
                        continue
                    
                    if idx % 25 == 0 and progress_callback:
                        progress_callback(f"Processed {results['items_processed']} items (skipped {skipped_items})")
                        
                except Exception as e:
                    skipped_items += 1
                    logger.debug(f"Skipped item {idx}: {e}")
                    continue
            
            results['success'] = results['items_processed'] > 0
            logger.info(f"Skip strategy processed {results['items_processed']} items, skipped {skipped_items}")
            
        except Exception as e:
            logger.error(f"Item-by-Item skip strategy failed: {e}")
            raise
        finally:
            if outlook:
                try:
                    outlook.detach_pst(Path(source_path))
                    for target_path in results['files_created']:
                        try:
                            outlook.detach_pst(Path(target_path))
                        except:
                            pass
                except:
                    pass
        
        return results
    
    def cleanup(self):
        """Clean up temporary resources"""
        try:
            if os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
        except Exception as e:
            logger.warning(f"Failed to cleanup temp directory: {e}")


# Global function for easy access
def handle_critically_full_pst(source_path: str, output_dir: str, split_by: str = "year", progress_callback=None):
    """
    Handle critically full PST using practical alternatives
    """
    handler = UltimatePSTHandler()
    try:
        return handler.handle_critically_full_pst(source_path, output_dir, split_by, progress_callback)
    finally:
        handler.cleanup()
