import os
import logging
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
            
            # Parse the text to extract name, role, location, and bullets
            name, role, location, experience_bullets = self._parse_cv_text(all_text, consultant_name)
            
            return {
                'name': name,
                'role': role,
                'location': location,
                'experience_bullets': experience_bullets,
                'headshot_image': headshot_image
            }
            
        except Exception as e:
            logger.error(f"Error extracting data from {cv_filepath}: {str(e)}")
            raise
    
    def _parse_cv_text(self, text_blocks: List[str], consultant_name: str) -> Tuple[str, str, str, List[str]]:
        """
        Parse text blocks to extract name, role, location, and experience bullets
        """
        name = consultant_name  # Use provided name as fallback
        role = ""
        location = ""
        experience_bullets = []
        
        for text_block in text_blocks:
            lines = [line.strip() for line in text_block.split('\n') if line.strip()]
            
            for line in lines:
                # Skip empty lines
                if not line:
                    continue
                    
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
            
        logger.info(f"Parsed data - Name: {name}, Role: {role}, Location: {location}, Bullets: {len(experience_bullets)}")
        
        return name, role, location, experience_bullets
    
    def generate_team_slide(self, consultant_names: List[str], filenames: List[str]) -> str:
        """
        Generate a team slide with 2x2 layout from consultant data
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
        
        # Define layout parameters for 2x2 grid
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        
        # Grid parameters
        margin = Inches(0.5)
        quadrant_width = (slide_width - 3 * margin) / 2
        quadrant_height = (slide_height - 3 * margin) / 2
        
        # Define positions for each quadrant (top-left, top-right, bottom-left, bottom-right)
        positions = [
            (margin, margin),  # Top-left
            (margin + quadrant_width + margin, margin),  # Top-right
            (margin, margin + quadrant_height + margin),  # Bottom-left
            (margin + quadrant_width + margin, margin + quadrant_height + margin)  # Bottom-right
        ]
        
        # Add each consultant to their quadrant
        for i, (data, position) in enumerate(zip(consultants_data, positions)):
            self._add_consultant_to_slide(slide, data, position, quadrant_width, quadrant_height)
        
        # Save the presentation
        output_filename = "Team_Slide_Output.pptx"
        output_path = os.path.join(self.output_folder, output_filename)
        prs.save(output_path)
        
        logger.info(f"Team slide saved to {output_path}")
        return output_path
    
    def _add_consultant_to_slide(self, slide, consultant_data: Dict, position: Tuple, width, height):
        """
        Add a single consultant's information to a quadrant of the slide
        """
        left, top = position
        
        # Define layout within quadrant
        image_width = Inches(2)
        image_height = Inches(2.5)
        text_left = left + image_width + Inches(0.2)
        text_width = width - image_width - Inches(0.2)
        
        # Add headshot image if available
        if consultant_data['headshot_image']:
            try:
                # Save image temporarily
                with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as temp_img:
                    temp_img.write(consultant_data['headshot_image'])
                    temp_img_path = temp_img.name
                
                # Add image to slide
                slide.shapes.add_picture(temp_img_path, left, top, image_width, image_height)
                
                # Clean up temp file
                os.unlink(temp_img_path)
                
            except Exception as e:
                logger.warning(f"Could not add image for {consultant_data['name']}: {str(e)}")
        
        # Add text content
        text_top = top
        
        # Name (bold)
        name_box = slide.shapes.add_textbox(text_left, text_top, text_width, Inches(0.4))
        name_frame = name_box.text_frame
        name_frame.margin_left = Pt(0)
        name_frame.margin_top = Pt(0)
        name_para = name_frame.paragraphs[0]
        name_run = name_para.add_run()
        name_run.text = consultant_data['name']
        name_run.font.bold = True
        name_run.font.size = Pt(14)
        name_run.font.color.rgb = RGBColor(0, 0, 0)
        
        # Role and location (italic)
        role_top = text_top + Inches(0.4)
        role_box = slide.shapes.add_textbox(text_left, role_top, text_width, Inches(0.6))
        role_frame = role_box.text_frame
        role_frame.margin_left = Pt(0)
        role_frame.margin_top = Pt(0)
        role_para = role_frame.paragraphs[0]
        role_run = role_para.add_run()
        role_run.text = f"{consultant_data['role']}\n{consultant_data['location']}"
        role_run.font.italic = True
        role_run.font.size = Pt(10)
        role_run.font.color.rgb = RGBColor(64, 64, 64)
        
        # Experience bullets
        bullets_top = role_top + Inches(0.6)
        bullets_height = height - Inches(1.0)  # Remaining space
        bullets_box = slide.shapes.add_textbox(text_left, bullets_top, text_width, bullets_height)
        bullets_frame = bullets_box.text_frame
        bullets_frame.margin_left = Pt(0)
        bullets_frame.margin_top = Pt(0)
        
        # Add each bullet point
        for i, bullet in enumerate(consultant_data['experience_bullets']):
            if i == 0:
                para = bullets_frame.paragraphs[0]
            else:
                para = bullets_frame.add_paragraph()
            
            para.level = 0
            run = para.add_run()
            run.text = f"• {bullet}"
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(32, 32, 32)
        
        logger.info(f"Added {consultant_data['name']} to slide at position {position}")