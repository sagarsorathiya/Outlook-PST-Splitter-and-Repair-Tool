"""Log export functionality for PST Splitter analysis."""
import logging
import json
import csv
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Optional
import zipfile
import tempfile


class PSTSplitterLogExporter:
    """Handles exporting logs and analysis data for troubleshooting."""
    
    def __init__(self):
        self.session_data = {
            'start_time': datetime.now().isoformat(),
            'groups_created': [],
            'errors': [],
            'performance_metrics': {},
            'space_liberation_attempts': [],
            'settings': {},
            'item_processing_log': []
        }
    
    def log_session_start(self, source_pst: str, output_dir: str, mode: str, **kwargs):
        """Log session start details."""
        self.session_data['settings'] = {
            'source_pst': source_pst,
            'output_dir': output_dir,
            'mode': mode,
            'timestamp': datetime.now().isoformat(),
            **kwargs
        }
        logging.info(f"📊 Session started: {mode} mode for {source_pst}")
    
    def log_group_creation(self, group_name: str, item_count: int, estimated_size: int = 0):
        """Log when a group/PST file is created."""
        group_info = {
            'name': group_name,
            'item_count': item_count,
            'estimated_size_mb': estimated_size / (1024*1024) if estimated_size else 0,
            'created_at': datetime.now().isoformat()
        }
        self.session_data['groups_created'].append(group_info)
        logging.info(f"📁 Group created: {group_name} ({item_count} items)")
    
    def log_error(self, error_type: str, message: str, context: Optional[Dict] = None):
        """Log an error with context."""
        error_info = {
            'type': error_type,
            'message': message,
            'context': context or {},
            'timestamp': datetime.now().isoformat()
        }
        self.session_data['errors'].append(error_info)
        logging.error(f"❌ {error_type}: {message}")
    
    def log_space_liberation_attempt(self, freed_mb: float, target_pst: str):
        """Log space liberation attempts to detect infinite loops."""
        attempt = {
            'freed_mb': freed_mb,
            'target_pst': target_pst,
            'timestamp': datetime.now().isoformat()
        }
        self.session_data['space_liberation_attempts'].append(attempt)
        logging.warning(f"🔄 Space liberation: {freed_mb:.1f}MB freed from {target_pst}")
    
    def log_performance_metric(self, metric_name: str, value: Any, unit: str = ""):
        """Log performance metrics."""
        self.session_data['performance_metrics'][metric_name] = {
            'value': value,
            'unit': unit,
            'timestamp': datetime.now().isoformat()
        }
        logging.info(f"⚡ Performance: {metric_name} = {value} {unit}")
    
    def log_item_processing(self, item_id: str, year: str, target_group: str, 
                           processing_time_ms: float, success: bool, error: Optional[str] = None):
        """Log detailed item processing information."""
        item_log = {
            'item_id': item_id[:50],  # Truncate for privacy
            'year': year,
            'target_group': target_group,
            'processing_time_ms': processing_time_ms,
            'success': success,
            'error': error,
            'timestamp': datetime.now().isoformat()
        }
        self.session_data['item_processing_log'].append(item_log)
    
    def export_analysis_report(self, output_path: Path) -> Path:
        """Export comprehensive analysis report."""
        self.session_data['end_time'] = datetime.now().isoformat()
        
        # Create analysis directory
        analysis_dir = output_path / f"PST_Analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        analysis_dir.mkdir(exist_ok=True)
        
        # Export JSON summary
        json_file = analysis_dir / "session_summary.json"
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(self.session_data, f, indent=2, ensure_ascii=False)
        
        # Export CSV for group analysis
        csv_file = analysis_dir / "groups_analysis.csv"
        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Group Name', 'Item Count', 'Estimated Size (MB)', 'Created At'])
            for group in self.session_data['groups_created']:
                writer.writerow([
                    group['name'], 
                    group['item_count'], 
                    round(group['estimated_size_mb'], 2),
                    group['created_at']
                ])
        
        # Export error analysis
        if self.session_data['errors']:
            error_file = analysis_dir / "errors_analysis.csv"
            with open(error_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Error Type', 'Message', 'Timestamp', 'Context'])
                for error in self.session_data['errors']:
                    writer.writerow([
                        error['type'],
                        error['message'],
                        error['timestamp'],
                        str(error.get('context', ''))
                    ])
        
        # Export space liberation analysis
        if self.session_data['space_liberation_attempts']:
            space_file = analysis_dir / "space_liberation_analysis.csv"
            with open(space_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Freed MB', 'Target PST', 'Timestamp'])
                for attempt in self.session_data['space_liberation_attempts']:
                    writer.writerow([
                        attempt['freed_mb'],
                        attempt['target_pst'],
                        attempt['timestamp']
                    ])
        
        # Export item processing log (sample for large datasets)
        if self.session_data['item_processing_log']:
            items_file = analysis_dir / "item_processing_sample.csv"
            with open(items_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Item ID', 'Year', 'Target Group', 'Processing Time (ms)', 'Success', 'Error'])
                # Export sample or all items if under 10000
                items_to_export = self.session_data['item_processing_log']
                if len(items_to_export) > 10000:
                    # Export every nth item to get representative sample
                    step = len(items_to_export) // 10000
                    items_to_export = items_to_export[::step]
                
                for item in items_to_export:
                    writer.writerow([
                        item['item_id'],
                        item['year'],
                        item['target_group'],
                        round(item['processing_time_ms'], 2),
                        item['success'],
                        item.get('error', '')
                    ])
        
        # Create performance summary
        perf_file = analysis_dir / "performance_summary.txt"
        with open(perf_file, 'w', encoding='utf-8') as f:
            f.write("PST Splitter Performance Summary\n")
            f.write("=" * 40 + "\n\n")
            
            f.write(f"Session Start: {self.session_data['start_time']}\n")
            f.write(f"Session End: {self.session_data.get('end_time', 'In Progress')}\n\n")
            
            f.write(f"Groups Created: {len(self.session_data['groups_created'])}\n")
            f.write(f"Errors Encountered: {len(self.session_data['errors'])}\n")
            f.write(f"Space Liberation Attempts: {len(self.session_data['space_liberation_attempts'])}\n\n")
            
            f.write("Performance Metrics:\n")
            for metric, data in self.session_data['performance_metrics'].items():
                f.write(f"  {metric}: {data['value']} {data['unit']}\n")
            
            if self.session_data['space_liberation_attempts']:
                f.write("\nSpace Liberation Analysis:\n")
                total_freed = sum(a['freed_mb'] for a in self.session_data['space_liberation_attempts'])
                f.write(f"  Total space freed: {total_freed:.1f} MB\n")
                f.write(f"  Average per attempt: {total_freed/len(self.session_data['space_liberation_attempts']):.1f} MB\n")
                
                # Check for potential infinite loops
                if len(self.session_data['space_liberation_attempts']) >= 3:
                    f.write("  ⚠️ WARNING: Multiple space liberation attempts detected!\n")
                    f.write("  This may indicate infinite loop conditions.\n")
        
        # Create README for the analysis
        readme_file = analysis_dir / "README.txt"
        with open(readme_file, 'w', encoding='utf-8') as f:
            f.write("PST Splitter Analysis Report\n")
            f.write("=" * 30 + "\n\n")
            f.write("This directory contains detailed analysis of your PST splitting session.\n\n")
            f.write("Files included:\n")
            f.write("- session_summary.json: Complete session data in JSON format\n")
            f.write("- groups_analysis.csv: Details of created PST groups\n")
            f.write("- performance_summary.txt: Human-readable performance summary\n")
            if self.session_data['errors']:
                f.write("- errors_analysis.csv: Detailed error analysis\n")
            if self.session_data['space_liberation_attempts']:
                f.write("- space_liberation_analysis.csv: Space management attempts\n")
            if self.session_data['item_processing_log']:
                f.write("- item_processing_sample.csv: Sample of item processing details\n")
            f.write("\nUse this data to:\n")
            f.write("- Understand why certain groups were created\n")
            f.write("- Analyze performance bottlenecks\n")
            f.write("- Troubleshoot errors or unexpected behavior\n")
            f.write("- Optimize future PST splitting operations\n")
        
        logging.info(f"📊 Analysis report exported to: {analysis_dir}")
        return analysis_dir

# Global instance for easy access
log_exporter = PSTSplitterLogExporter()
