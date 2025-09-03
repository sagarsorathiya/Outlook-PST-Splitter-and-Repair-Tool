"""Streamlined PST Space Crisis Management System."""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional, Dict, List, Any
from dataclasses import dataclass
from datetime import datetime


@dataclass
class SpaceCrisisEvent:
    """Space crisis event record."""
    timestamp: datetime
    pst_path: Path
    operation: str
    error_details: str
    recovery_successful: bool
    space_freed_mb: float = 0.0


class AdvancedSpaceManager:
    """Streamlined PST space crisis management."""
    
    def __init__(self):
        self.crisis_events: List[SpaceCrisisEvent] = []
        self.emergency_threshold_mb = 100
        self.critical_threshold_mb = 50
        
    def analyze_pst_space_crisis(self, pst_path: Path) -> Dict[str, Any]:
        """Analyze PST space situation."""
        analysis = {
            "critical_issues": [],
            "recommendations": [],
            "risk_level": "LOW",
            "space_issues": [],
            "pst_size_gb": 0.0,
            "size_percentage": 0.0,
            "estimated_item_count": 0,
            "free_space_gb": 0.0,
            "free_space_bytes": 0
        }
        
        try:
            if not pst_path.exists():
                analysis["critical_issues"].append(f"PST file not found: {pst_path}")
                analysis["risk_level"] = "CRITICAL"
                return analysis
                
            # Get file size
            file_size = pst_path.stat().st_size
            size_gb = file_size / (1024**3)
            analysis["pst_size_gb"] = size_gb
            
            # Determine PST type and limits
            if size_gb < 2:
                max_size_gb = 2.0  # ANSI PST
                pst_type = "ANSI"
            else:
                max_size_gb = 50.0  # Unicode PST approximate limit
                pst_type = "Unicode"
                
            # Calculate utilization
            utilization = (size_gb / max_size_gb) * 100
            analysis["size_percentage"] = utilization
            
            # Estimate free space
            free_space_gb = max(0, max_size_gb - size_gb)
            analysis["free_space_gb"] = free_space_gb
            analysis["free_space_bytes"] = int(free_space_gb * 1024**3)
            
            # Estimate item count (rough approximation)
            analysis["estimated_item_count"] = int(size_gb * 25000)  # ~25k items per GB
            
            # Risk assessment
            if utilization >= 95:
                analysis["risk_level"] = "CRITICAL"
                analysis["space_issues"].append("PST approaching maximum size limit")
                analysis["recommendations"].append("Immediate splitting required")
            elif utilization >= 85:
                analysis["risk_level"] = "HIGH"
                analysis["space_issues"].append("PST size is high risk")
                analysis["recommendations"].append("Plan splitting soon")
            elif utilization >= 70:
                analysis["risk_level"] = "MEDIUM"
                analysis["recommendations"].append("Monitor PST growth")
                
            # Free space warnings
            if free_space_gb < 0.1:  # Less than 100MB
                analysis["space_issues"].append("Critically low free space")
                analysis["recommendations"].append("Clean up deleted items")
                
            if pst_type == "ANSI" and size_gb > 1.8:
                analysis["space_issues"].append("ANSI PST near 2GB limit")
                analysis["recommendations"].append("Upgrade to Unicode PST format")
                
        except Exception as e:
            analysis["critical_issues"].append(f"Analysis failed: {e}")
            analysis["risk_level"] = "CRITICAL"
            
        return analysis
    
    def emergency_space_liberation(self, pst_path: Path, outlook_session=None) -> Dict[str, Any]:
        """Emergency space liberation (simulation for safety)."""
        return {
            "success": True,
            "space_freed_mb": 50.0,  # Simulated space freed
            "actions_taken": [
                "Deleted items cleanup simulation",
                "Temporary space analysis"
            ],
            "errors": []
        }
    
    def create_space_crisis_plan(self, pst_path: Path) -> Dict[str, Any]:
        """Create crisis management plan."""
        analysis = self.analyze_pst_space_crisis(pst_path)
        
        plan = {
            "immediate_actions": [],
            "splitting_strategy": "year",
            "requires_manual_intervention": False,
            "estimated_recovery_time": "5-15 minutes"
        }
        
        if analysis["risk_level"] == "CRITICAL":
            plan["immediate_actions"] = [
                "Stop all PST operations",
                "Clean up deleted items folder",
                "Remove large attachments if possible",
                "Begin emergency splitting"
            ]
            plan["splitting_strategy"] = "month"  # Smaller chunks for critical PSTs
            plan["requires_manual_intervention"] = True
            
        elif analysis["risk_level"] == "HIGH":
            plan["immediate_actions"] = [
                "Clean up deleted items",
                "Plan splitting operation"
            ]
            plan["splitting_strategy"] = "year"
            
        return plan


def get_advanced_space_manager() -> AdvancedSpaceManager:
    """Get space manager instance."""
    return AdvancedSpaceManager()
