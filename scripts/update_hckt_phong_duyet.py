"""
Update HCKT criteria to assign them to Phòng Hậu cần - Kỹ thuật for approval.
Run this script once to fix the missing phong_duyet assignment.

Usage:
    python scripts/update_hckt_phong_duyet.py
"""
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from app import create_app, db
from app.models.nomination import TieuChi
import json

def main():
    app = create_app()
    with app.app_context():
        # Find all HCKT criteria (ma_truong starts with 'hckt_')
        hckt_criteria = TieuChi.query.filter(TieuChi.ma_truong.like('hckt_%')).all()
        
        if not hckt_criteria:
            print("No HCKT criteria found")
            return
        
        print(f"Found {len(hckt_criteria)} HCKT criteria")
        
        phong_hckt = 'Phòng Hậu cần - Kỹ thuật'
        updated = 0
        
        for tc in hckt_criteria:
            # Get current phong_duyet list
            current = tc.phong_duyet or []
            
            # Add PHONG_HAUCANKYTHUAT if not already present
            if phong_hckt not in current:
                current.append(phong_hckt)
                tc.phong_duyet = current
                updated += 1
                print(f"  Updated: {tc.ma_truong}")
            else:
                print(f"  Skipped: {tc.ma_truong}")
        
        if updated > 0:
            db.session.commit()
            print(f"\nUpdated {updated} criteria successfully")
        else:
            print("\nAll criteria already assigned correctly")

if __name__ == '__main__':
    main()
