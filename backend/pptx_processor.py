import os
import logging
import re
from typing import List, Dict, Tuple, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from PIL import Image
import io
import tempfile

logger = logging.getLogger(__name__)

class PowerPointProcessor:
    def __init__(self, cvs_folder: str, output_folder: str, examples_folder: str):
        self.cvs_folder = cvs_folder
        self.output_folder = output_folder
        self.examples_folder = examples_folder
        
    def find_cv_file(self, consultant_name: str) -> Optional[str]:
        """
        Find CV file for a consultant name using flexible matching
        """
        # Get all .pptx files in the CVs folder
        try:
            cv_files = [f for f in os.listdir(self.cvs_folder) if f.endswith('.pptx')]
        except OSError:
            logger.error(f"Could not list files in {self.cvs_folder}")
            return None
        
        # First try exact match with standard naming rule
        standard_filename = consultant_name.replace(" ", "_").replace("-", "") + ".pptx"
        if standard_filename in cv_files:
            logger.info(f"Found exact match for {consultant_name}: {standard_filename}")
            return standard_filename
        
        # Try flexible matching - look for files that contain the consultant's name parts
        name_parts = consultant_name.lower().split()
        
        for cv_file in cv_files:
            # Remove .pptx extension and convert to lowercase for comparison
            file_base = cv_file[:-5].lower()  # Remove ".pptx"
            
            # Check if all name parts are present in the filename
            if all(part in file_base for part in name_parts):
                logger.info(f"Found fuzzy match for {consultant_name}: {cv_file}")
                return cv_file
        
        # If no match found, log available files for debugging
        logger.warning(f"No CV file found for {consultant_name}. Available files: {cv_files}")
        return None
        
    def extract_consultant_data(self, cv_filepath: str, consultant_name: str) -> Dict:
        """
        Extract consultant data from a CV PowerPoint file
        Returns: {
            'name': str,
            'role': str, 
            'location': str,
            'experience_bullets': List[str],
            'headshot_image': bytes or None
        }
        """
        logger.info(f"Extracting data from {cv_filepath}")
        
        try:
            prs = Presentation(cv_filepath)
            
            # Assume we're working with the first slide
            if len(prs.slides) == 0:
                raise ValueError(f"No slides found in {cv_filepath}")
                
            slide = prs.slides[0]
            
            # Extract headshot (first image shape)
            headshot_image = None
            for shape in slide.shapes:
                if hasattr(shape, 'image'):
                    # Found an image, extract it
                    image_stream = io.BytesIO(shape.image.blob)
                    headshot_image = image_stream.getvalue()
                    logger.info(f"Extracted headshot image from {cv_filepath}")
                    break
            
            # Extract text content
            all_text = []
            for shape in slide.shapes:
                if hasattr(shape, 'text') and shape.text.strip():
                    all_text.append(shape.text.strip())
            
            # Parse the text to extract name, role, location, bullets, and years of experience
            name, role, location, experience_bullets, years_experience = self._parse_cv_text(all_text, consultant_name)
            
            return {
                'name': name,
                'role': role,
                'location': location,
                'experience_bullets': experience_bullets,
                'years_experience': years_experience,
                'headshot_image': headshot_image
            }
            
        except Exception as e:
            logger.error(f"Error extracting data from {cv_filepath}: {str(e)}")
            raise
    
    def _parse_cv_text(self, text_blocks: List[str], consultant_name: str) -> Tuple[str, str, str, List[str], str]:
        """
        Parse text blocks to extract name, role, location, experience bullets, and years of experience
        """
        name = consultant_name  # Use provided name as fallback
        role = ""
        location = ""
        experience_bullets = []
        years_experience = ""
        
        for text_block in text_blocks:
            lines = [line.strip() for line in text_block.split('\n') if line.strip()]
            
            for line in lines:
                # Skip empty lines
                if not line:
                    continue
                    
                # Look for years of experience patterns
                if not years_experience:
                    # Pattern like "3+ years", "2 years", "5+ years of consulting"
                    years_pattern = re.search(r'(\d+)\+?\s*years?\s*(of\s*)?(consulting|experience|industry)', line.lower())
                    if years_pattern:
                        years_num = years_pattern.group(1)
                        years_experience = f"{years_num}+ years of consulting and\nindustry experience"
                    
                # Look for bullet points (lines starting with •, -, or similar)
                if any(line.startswith(bullet) for bullet in ['•', '-', '▪', '◦', '→']):
                    clean_bullet = line.lstrip('•-▪◦→ ').strip()
                    if clean_bullet and len(clean_bullet) > 10:  # Only meaningful bullets
                        experience_bullets.append(clean_bullet)
                
                # Look for role/location patterns
                elif any(keyword in line.lower() for keyword in ['consultant', 'manager', 'director', 'analyst', 'partner']):
                    if not role and len(line) < 100:  # Reasonable length for a role
                        role = line
                        
                elif any(keyword in line.lower() for keyword in ['london', 'new york', 'paris', 'berlin', 'zurich', 'geneva', 'munich']):
                    if not location and len(line) < 50:  # Reasonable length for location
                        location = line
        
        # Limit experience bullets to 4 most relevant ones
        experience_bullets = experience_bullets[:4]
        
        # Fallback values if not found
        if not role:
            role = "Consultant"
        if not location:
            location = "Location"
        if not years_experience:
            years_experience = "3+ years of consulting and\nindustry experience"  # Default fallback
            
        logger.info(f"Parsed data - Name: {name}, Role: {role}, Location: {location}, Years: {years_experience}, Bullets: {len(experience_bullets)}")
        
        return name, role, location, experience_bullets, years_experience
    
    def generate_team_slide(self, consultant_names: List[str], filenames: List[str]) -> str:
        """
        Generate a team slide matching the exact format of the reference example
        Returns path to the generated PowerPoint file
        """
        logger.info("Starting team slide generation")
        
        # Extract data from all CV files
        consultants_data = []
        for name, filename in zip(consultant_names, filenames):
            cv_filepath = os.path.join(self.cvs_folder, filename)
            
            if not os.path.exists(cv_filepath):
                raise FileNotFoundError(f"CV file not found: {filename}")
                
            data = self.extract_consultant_data(cv_filepath, name)
            consultants_data.append(data)
        
        # Create new presentation
        prs = Presentation()
        
        # Set slide size to 16:9 (standard)
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        
        # Add a slide
        slide_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(slide_layout)
        
        # Define layout parameters matching the reference
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        
        # Add title "Your Kearney project team" in purple on the left
        title_left = Inches(0.5)
        title_top = Inches(0.8)
        title_width = Inches(3)
        title_height = Inches(1.5)
        
        title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
        title_frame = title_box.text_frame
        title_frame.margin_left = Pt(0)
        title_frame.margin_top = Pt(0)
        title_para = title_frame.paragraphs[0]
        title_run = title_para.add_run()
        title_run.text = "Your Kearney\nproject team"
        title_run.font.size = Pt(24)
        title_run.font.color.rgb = RGBColor(102, 45, 145)  # Consulting Purple
        title_run.font.bold = True
        
        # Define horizontal layout for consultants
        consultant_start_left = Inches(4.2)
        consultant_width = Inches(2.2)
        consultant_spacing = Inches(0.1)
        
        # Add each consultant horizontally
        for i, data in enumerate(consultants_data):
            consultant_left = consultant_start_left + i * (consultant_width + consultant_spacing)
            self._add_consultant_to_slide(slide, data, consultant_left, consultant_width)
        
        # Add "Project Team" label at bottom left matching reference
        project_label_left = Inches(0.5)
        project_label_top = Inches(6.8)
        project_label_width = Inches(1.5)
        project_label_height = Inches(0.4)
        
        project_box = slide.shapes.add_textbox(project_label_left, project_label_top, project_label_width, project_label_height)
        project_frame = project_box.text_frame
        project_frame.margin_left = Pt(0)
        project_frame.margin_top = Pt(0)
        project_para = project_frame.paragraphs[0]
        project_run = project_para.add_run()
        project_run.text = "Project Team"
        project_run.font.size = Pt(10)
        project_run.font.color.rgb = RGBColor(64, 64, 64)
        
        # Save the presentation
        output_filename = "Team_Slide_Output.pptx"
        output_path = os.path.join(self.output_folder, output_filename)
        prs.save(output_path)
        
        logger.info(f"Team slide saved to {output_path}")
        return output_path
    
    def _add_consultant_to_slide(self, slide, consultant_data: Dict, left_position, width):
        """
        Add a single consultant's information in vertical layout matching the reference
        """
        # Define vertical layout positions
        photo_top = Inches(0.8)
        photo_width = Inches(1.8)
        photo_height = Inches(2.2)
        
        # Center the photo within the column width
        photo_left = left_position + (width - photo_width) / 2
        
        # Add headshot image if available
        if consultant_data['headshot_image']:
            try:
                # Save image temporarily
                with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as temp_img:
                    temp_img.write(consultant_data['headshot_image'])
                    temp_img_path = temp_img.name
                
                # Add image to slide with proper cropping/resizing
                slide.shapes.add_picture(temp_img_path, photo_left, photo_top, photo_width, photo_height)
                
                # Clean up temp file
                os.unlink(temp_img_path)
                
            except Exception as e:
                logger.warning(f"Could not add image for {consultant_data['name']}: {str(e)}")
        
        # Name in purple (positioned under photo)
        name_top = photo_top + photo_height + Inches(0.15)
        name_box = slide.shapes.add_textbox(left_position, name_top, width, Inches(0.4))
        name_frame = name_box.text_frame
        name_frame.margin_left = Pt(0)
        name_frame.margin_top = Pt(0)
        name_para = name_frame.paragraphs[0]
        name_para.alignment = PP_ALIGN.CENTER
        name_run = name_para.add_run()
        name_run.text = consultant_data['name']
        name_run.font.bold = True
        name_run.font.size = Pt(12)
        name_run.font.color.rgb = RGBColor(102, 45, 145)  # Consulting Purple
        
        # Role and location summary (e.g., "Sr Consultant, Munich")
        summary_top = name_top + Inches(0.4)
        summary_box = slide.shapes.add_textbox(left_position, summary_top, width, Inches(0.3))
        summary_frame = summary_box.text_frame
        summary_frame.margin_left = Pt(0)
        summary_frame.margin_top = Pt(0)
        summary_para = summary_frame.paragraphs[0]
        summary_para.alignment = PP_ALIGN.CENTER
        summary_run = summary_para.add_run()
        
        # Parse role to create summary line like "Sr Consultant, Munich"
        role_parts = consultant_data['role'].split()
        if 'consultant' in consultant_data['role'].lower():
            if any(word in consultant_data['role'].lower() for word in ['senior', 'sr', 'lead']):
                summary_text = "Sr Consultant"
            else:
                summary_text = "Consultant"
        else:
            summary_text = consultant_data['role']
        
        # Add location if available
        if consultant_data['location'] and consultant_data['location'] != "Location":
            summary_text += f", {consultant_data['location']}"
        
        summary_run.text = summary_text
        summary_run.font.size = Pt(10)
        summary_run.font.color.rgb = RGBColor(64, 64, 64)
        
        # Experience summary line (e.g., "3+ years of consulting and industry experience")
        exp_summary_top = summary_top + Inches(0.35)
        exp_summary_box = slide.shapes.add_textbox(left_position, exp_summary_top, width, Inches(0.3))
        exp_summary_frame = exp_summary_box.text_frame
        exp_summary_frame.margin_left = Pt(0)
        exp_summary_frame.margin_top = Pt(0)
        exp_summary_para = exp_summary_frame.paragraphs[0]
        exp_summary_run = exp_summary_para.add_run()
        exp_summary_run.text = consultant_data['years_experience']
        exp_summary_run.font.size = Pt(9)
        exp_summary_run.font.color.rgb = RGBColor(64, 64, 64)
        
        # Experience bullets (limit to 3 as per requirements)
        bullets_top = exp_summary_top + Inches(0.6)
        bullets_box = slide.shapes.add_textbox(left_position, bullets_top, width, Inches(3.5))
        bullets_frame = bullets_box.text_frame
        bullets_frame.margin_left = Pt(0)
        bullets_frame.margin_top = Pt(0)
        
        # Add up to 3 bullet points
        experience_bullets = consultant_data['experience_bullets'][:3]
        
        for i, bullet in enumerate(experience_bullets):
            if i == 0:
                para = bullets_frame.paragraphs[0]
            else:
                para = bullets_frame.add_paragraph()
            
            para.level = 0
            run = para.add_run()
            run.text = f"– {bullet}"  # Use en-dash as in reference
            run.font.size = Pt(8)
            run.font.color.rgb = RGBColor(32, 32, 32)
            para.space_after = Pt(6)  # Add spacing between bullets
        
        logger.info(f"Added {consultant_data['name']} to slide at position {left_position}")