#!/usr/bin/env python3
"""
Azure Entra Group Member Comparison Pipeline - Azure DevOps Version

This script:
1. Fetches current group members from Azure Entra ID
2. Compares with previously stored member list
3. Generates a PDF report highlighting changes (new members in green, removed in amber)

Prerequisites:
- pip install azure-identity msal requests reportlab
- Azure app registration with appropriate permissions
- Group.Read.All permission in Azure AD
"""

import json
import os
import datetime
import sys
from typing import Dict, List, Set, Tuple
from dataclasses import dataclass
from pathlib import Path

import requests
from azure.identity import ClientSecretCredential
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch


@dataclass
class GroupMember:
    """Represents a group member with essential information"""
    id: str
    display_name: str
    user_principal_name: str
    mail: str = None
    
    def to_dict(self) -> Dict:
        return {
            'id': self.id,
            'display_name': self.display_name,
            'user_principal_name': self.user_principal_name,
            'mail': self.mail
        }
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'GroupMember':
        return cls(
            id=data['id'],
            display_name=data['display_name'],
            user_principal_name=data['user_principal_name'],
            mail=data.get('mail')
        )


class EntraGroupManager:
    """Manages Azure Entra ID group operations"""
    
    def __init__(self, tenant_id: str, client_id: str, client_secret: str):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.credential = ClientSecretCredential(tenant_id, client_id, client_secret)
        self.graph_url = "https://graph.microsoft.com/v1.0"
        
    def get_access_token(self) -> str:
        """Get access token for Microsoft Graph API"""
        token = self.credential.get_token("https://graph.microsoft.com/.default")
        return token.token
    
    def get_group_members(self, group_id: str) -> List[GroupMember]:
        """Fetch current group members from Azure Entra ID"""
        headers = {
            'Authorization': f'Bearer {self.get_access_token()}',
            'Content-Type': 'application/json'
        }
        
        members = []
        url = f"{self.graph_url}/groups/{group_id}/members"
        
        while url:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            
            data = response.json()
            
            for member in data.get('value', []):
                if member.get('@odata.type') == '#microsoft.graph.user':
                    members.append(GroupMember(
                        id=member['id'],
                        display_name=member.get('displayName', ''),
                        user_principal_name=member.get('userPrincipalName', ''),
                        mail=member.get('mail')
                    ))
            
            url = data.get('@odata.nextLink')
        
        return members


class MembershipComparator:
    """Compares group memberships and identifies changes"""
    
    @staticmethod
    def compare_memberships(
        current_members: List[GroupMember], 
        previous_members: List[GroupMember]
    ) -> Tuple[List[GroupMember], List[GroupMember], List[GroupMember]]:
        """
        Compare current and previous member lists
        
        Returns:
            Tuple of (new_members, removed_members, unchanged_members)
        """
        current_ids = {member.id for member in current_members}
        previous_ids = {member.id for member in previous_members}
        
        # Create lookup dictionaries
        current_dict = {member.id: member for member in current_members}
        previous_dict = {member.id: member for member in previous_members}
        
        # Find changes
        new_member_ids = current_ids - previous_ids
        removed_member_ids = previous_ids - current_ids
        unchanged_member_ids = current_ids & previous_ids
        
        new_members = [current_dict[id] for id in new_member_ids]
        removed_members = [previous_dict[id] for id in removed_member_ids]
        unchanged_members = [current_dict[id] for id in unchanged_member_ids]
        
        return new_members, removed_members, unchanged_members


class PDFReportGenerator:
    """Generates PDF reports for membership changes"""
    
    def __init__(self, output_path: str = "membership_report.pdf"):
        self.output_path = output_path
        self.styles = getSampleStyleSheet()
        
        # Define custom styles
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=18,
            textColor=colors.darkblue,
            spaceAfter=30
        )
        
        self.section_style = ParagraphStyle(
            'CustomSection',
            parent=self.styles['Heading2'],
            fontSize=14,
            textColor=colors.black,
            spaceAfter=12
        )
    
    def generate_report(
        self, 
        group_name: str,
        new_members: List[GroupMember], 
        removed_members: List[GroupMember], 
        unchanged_members: List[GroupMember],
        comparison_date: datetime.datetime
    ):
        """Generate PDF report with membership changes"""
        
        doc = SimpleDocTemplate(
            self.output_path,
            pagesize=A4,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=18
        )
        
        story = []
        
        # Title
        title = Paragraph(f"Group Membership Report: {group_name}", self.title_style)
        story.append(title)
        
        # Report metadata
        metadata = Paragraph(
            f"<b>Report Date:</b> {comparison_date.strftime('%Y-%m-%d %H:%M:%S')}<br/>"
            f"<b>Total Current Members:</b> {len(new_members) + len(unchanged_members)}<br/>"
            f"<b>New Members:</b> {len(new_members)}<br/>"
            f"<b>Removed Members:</b> {len(removed_members)}<br/>",
            self.styles['Normal']
        )
        story.append(metadata)
        story.append(Spacer(1, 20))
        
        # Summary section
        if new_members or removed_members:
            story.append(Paragraph("Summary of Changes", self.section_style))
            
            # New members section
            if new_members:
                story.append(Paragraph("New Members", self.styles['Heading3']))
                new_data = [['Display Name', 'User Principal Name', 'Email']]
                for member in new_members:
                    new_data.append([
                        member.display_name,
                        member.user_principal_name,
                        member.mail or 'N/A'
                    ])
                
                new_table = Table(new_data, colWidths=[2*inch, 2.5*inch, 2*inch])
                new_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.darkgreen),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 12),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.lightgreen),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                story.append(new_table)
                story.append(Spacer(1, 20))
            
            # Removed members section
            if removed_members:
                story.append(Paragraph("Removed Members", self.styles['Heading3']))
                removed_data = [['Display Name', 'User Principal Name', 'Email']]
                for member in removed_members:
                    removed_data.append([
                        member.display_name,
                        member.user_principal_name,
                        member.mail or 'N/A'
                    ])
                
                removed_table = Table(removed_data, colWidths=[2*inch, 2.5*inch, 2*inch])
                removed_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.darkorange),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 12),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.bisque),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                story.append(removed_table)
                story.append(Spacer(1, 20))
        else:
            story.append(Paragraph("No membership changes detected.", self.styles['Normal']))
            story.append(Spacer(1, 20))
        
        # Complete member list
        if unchanged_members:
            story.append(Paragraph("Current Members (Unchanged)", self.section_style))
            unchanged_data = [['Display Name', 'User Principal Name', 'Email']]
            for member in sorted(unchanged_members, key=lambda x: x.display_name.lower()):
                unchanged_data.append([
                    member.display_name,
                    member.user_principal_name,
                    member.mail or 'N/A'
                ])
            
            unchanged_table = Table(unchanged_data, colWidths=[2*inch, 2.5*inch, 2*inch])
            unchanged_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
            ]))
            story.append(unchanged_table)
        
        doc.build(story)
        print(f"PDF report generated: {self.output_path}")


class GroupMembershipPipeline:
    """Main pipeline orchestrator - Azure DevOps version"""
    
    def __init__(self):
        self.data_file = "previous_members.json"
        self.reports_dir = "reports"
        self.load_config_from_env()
        
        # Ensure reports directory exists
        Path(self.reports_dir).mkdir(exist_ok=True)
        
        self.entra_manager = EntraGroupManager(
            self.config['tenant_id'],
            self.config['client_id'],
            self.config['client_secret']
        )
    
    def load_config_from_env(self):
        """Load configuration from environment variables (Azure DevOps)"""
        required_vars = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET', 'GROUP_ID', 'GROUP_NAME']
        
        self.config = {}
        missing_vars = []
        
        for var in required_vars:
            value = os.environ.get(var)
            if not value:
                missing_vars.append(var)
            else:
                self.config[var.lower()] = value
        
        if missing_vars:
            raise ValueError(f"Missing required environment variables: {', '.join(missing_vars)}")
        
        print(f"Configuration loaded for group: {self.config['group_name']}")
    
    def load_previous_members(self) -> List[GroupMember]:
        """Load previously stored member list"""
        try:
            with open(self.data_file, 'r') as f:
                data = json.load(f)
                print(f"Loaded previous member data from {data.get('last_updated', 'unknown date')}")
                return [GroupMember.from_dict(member) for member in data['members']]
        except FileNotFoundError:
            print("No previous member data found. This will be treated as initial run.")
            return []
    
    def save_current_members(self, members: List[GroupMember]):
        """Save current member list for future comparison"""
        data = {
            'last_updated': datetime.datetime.now().isoformat(),
            'members': [member.to_dict() for member in members]
        }
        with open(self.data_file, 'w') as f:
            json.dump(data, f, indent=2)
        print(f"Current member list saved to {self.data_file}")
    
    def run(self):
        """Execute the complete pipeline"""
        try:
            print("=" * 60)
            print("Starting Azure Entra Group Membership Comparison Pipeline")
            print("=" * 60)
            
            # Fetch current members
            print(f"\n1. Fetching current members for group: {self.config['group_id']}")
            current_members = self.entra_manager.get_group_members(self.config['group_id'])
            print(f"   ✓ Found {len(current_members)} current members")
            
            # Load previous members
            print(f"\n2. Loading previous member data...")
            previous_members = self.load_previous_members()
            print(f"   ✓ Loaded {len(previous_members)} previous members")
            
            # Compare memberships
            print(f"\n3. Comparing memberships...")
            comparator = MembershipComparator()
            new_members, removed_members, unchanged_members = comparator.compare_memberships(
                current_members, previous_members
            )
            
            print(f"   ✓ Comparison results:")
            print(f"     - New members: {len(new_members)}")
            print(f"     - Removed members: {len(removed_members)}")
            print(f"     - Unchanged members: {len(unchanged_members)}")
            
            # Set Azure DevOps pipeline variables for downstream tasks
            changes_detected = len(new_members) > 0 or len(removed_members) > 0
            print(f"\n##vso[task.setvariable variable=ChangesDetected;isOutput=true]{str(changes_detected).lower()}")
            print(f"##vso[task.setvariable variable=NewMembersCount;isOutput=true]{len(new_members)}")
            print(f"##vso[task.setvariable variable=RemovedMembersCount;isOutput=true]{len(removed_members)}")
            
            # Generate PDF report
            print(f"\n4. Generating PDF report...")
            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            report_filename = os.path.join(self.reports_dir, f"membership_report_{timestamp}.pdf")
            
            pdf_generator = PDFReportGenerator(report_filename)
            pdf_generator.generate_report(
                self.config['group_name'],
                new_members,
                removed_members,
                unchanged_members,
                datetime.datetime.now()
            )
            print(f"   ✓ PDF report generated: {report_filename}")
            
            # Save current members for next comparison
            print(f"\n5. Saving current member data for next comparison...")
            self.save_current_members(current_members)
            print(f"   ✓ Current member list saved")
            
            # Summary
            print(f"\n" + "=" * 60)
            print("PIPELINE SUMMARY")
            print("=" * 60)
            print(f"Group: {self.config['group_name']}")
            print(f"Total Members: {len(current_members)}")
            print(f"Changes Detected: {'Yes' if changes_detected else 'No'}")
            if changes_detected:
                print(f"New Members: {len(new_members)}")
                print(f"Removed Members: {len(removed_members)}")
                if new_members:
                    print("New Members List:")
                    for member in new_members:
                        print(f"  + {member.display_name} ({member.user_principal_name})")
                if removed_members:
                    print("Removed Members List:")
                    for member in removed_members:
                        print(f"  - {member.display_name} ({member.user_principal_name})")
            print(f"Report: {report_filename}")
            print("=" * 60)
            
            print("\n✅ Pipeline completed successfully!")
            
            # Exit with specific code based on changes
            return 0 if not changes_detected else 1
            
        except Exception as e:
            print(f"\n❌ Pipeline failed with error: {str(e)}")
            print(f"##vso[task.logissue type=error]Pipeline failed: {str(e)}")
            raise


def main():
    """Main entry point"""
    try:
        pipeline = GroupMembershipPipeline()
        exit_code = pipeline.run()
        sys.exit(exit_code)
    except Exception as e:
        print(f"Fatal error: {str(e)}")
        sys.exit(2)


if __name__ == "__main__":
    main()