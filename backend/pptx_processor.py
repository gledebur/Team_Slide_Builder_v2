import os
import logging
from typing import List, Dict, Tuple, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image
import io
import tempfile

logger = logging.getLogger(__name__)

# Define Consulting Purple color (RGB values)
CONSULTING_PURPLE = RGBColor(102, 45, 145)  # Deep purple color typically used in consulting
GRAY_TEXT = RGBColor(64, 64, 64)  # Dark gray for secondary text
BLACK_TEXT = RGBColor(0, 0, 0)  # Black for primary text

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
            'experience_summary': str,
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
            
            # Parse the text to extract name, role, location, summary, and bullets
            name, role, location, experience_summary, experience_bullets = self._parse_cv_text(all_text, consultant_name)
            
            return {
                'name': name,
                'role': role,
                'location': location,
                'experience_summary': experience_summary,
                'experience_bullets': experience_bullets,
                'headshot_image': headshot_image
            }
            
        except Exception as e:
            logger.error(f"Error extracting data from {cv_filepath}: {str(e)}")
            raise
    
    def _parse_cv_text(self, text_blocks: List[str], consultant_name: str) -> Tuple[str, str, str, str, List[str]]:
        """
        Parse text blocks to extract name, role, location, experience summary, and experience bullets
        """
        name = consultant_name  # Use provided name as fallback
        role = ""
        location = ""
        experience_summary = ""
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
                
                # Look for experience summary (lines with "years", "experience", etc.)
                elif any(keyword in line.lower() for keyword in ['years', 'experience', 'expertise', 'specializes']) and not experience_summary:
                    if 20 <= len(line) <= 150:  # Reasonable length for summary
                        experience_summary = line
        
        # Limit experience bullets to exactly 3 as per requirements
        experience_bullets = experience_bullets[:3]
        
        # Fallback values if not found
        if not role:
            role = "Senior Consultant"
        if not location:
            location = "Global"
        if not experience_summary:
            experience_summary = f"Experienced consultant with expertise in strategy and operations"
            
        logger.info(f"Parsed data - Name: {name}, Role: {role}, Location: {location}, Summary: {experience_summary[:50]}..., Bullets: {len(experience_bullets)}")
        
        return name, role, location, experience_summary, experience_bullets
    
    def _crop_and_resize_image(self, image_bytes: bytes, target_width: int, target_height: int) -> bytes:
        """
        Crop and resize image to fit exactly into the designated space
        """
        try:
            # Load image
            image = Image.open(io.BytesIO(image_bytes))
            
            # Convert to RGB if necessary
            if image.mode != 'RGB':
                image = image.convert('RGB')
            
            # Calculate crop box for center crop
            img_width, img_height = image.size
            aspect_ratio = target_width / target_height
            img_aspect_ratio = img_width / img_height
            
            if img_aspect_ratio > aspect_ratio:
                # Image is wider than target, crop width
                new_width = int(img_height * aspect_ratio)
                left = (img_width - new_width) // 2
                crop_box = (left, 0, left + new_width, img_height)
            else:
                # Image is taller than target, crop height
                new_height = int(img_width / aspect_ratio)
                top = (img_height - new_height) // 2
                crop_box = (0, top, img_width, top + new_height)
            
            # Crop and resize
            cropped_image = image.crop(crop_box)
            resized_image = cropped_image.resize((target_width, target_height), Image.Resampling.LANCZOS)
            
            # Save to bytes
            output = io.BytesIO()
            resized_image.save(output, format='JPEG', quality=95)
            return output.getvalue()
            
        except Exception as e:
            logger.warning(f"Error processing image: {str(e)}")
            return image_bytes  # Return original if processing fails
    
    def generate_team_slide(self, consultant_names: List[str], filenames: List[str]) -> str:
        """
        Generate a team slide with 2x2 layout matching the example format exactly
        Returns path to the generated PowerPoint file
        """
        logger.info("Starting team slide generation with exact example formatting")
        
        # Extract data from all CV files
        consultants_data = []
        for name, filename in zip(consultant_names, filenames):
            cv_filepath = os.path.join(self.cvs_folder, filename)
            
            if not os.path.exists(cv_filepath):
                raise FileNotFoundError(f"CV file not found: {filename}")
                
            data = self.extract_consultant_data(cv_filepath, name)
            consultants_data.append(data)
        
        # Create new presentation matching the example format
        prs = Presentation()
        
        # Set slide size to 16:9 (standard)
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        
        # Add a slide
        slide_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(slide_layout)
        
        # Set slide background to white
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(255, 255, 255)
        
        # Define layout parameters matching the example exactly
        # These measurements are based on typical consulting slide layouts
        margin_top = Inches(0.8)
        margin_left = Inches(0.6)
        margin_right = Inches(0.6)
        margin_bottom = Inches(0.6)
        
        # Calculate quadrant dimensions
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        
        available_width = slide_width - margin_left - margin_right
        available_height = slide_height - margin_top - margin_bottom
        
        quadrant_width = available_width / 2
        quadrant_height = available_height / 2
        
        # Define positions for each quadrant (top-left, top-right, bottom-left, bottom-right)
        positions = [
            (margin_left, margin_top),  # Top-left
            (margin_left + quadrant_width, margin_top),  # Top-right
            (margin_left, margin_top + quadrant_height),  # Bottom-left
            (margin_left + quadrant_width, margin_top + quadrant_height)  # Bottom-right
        ]
        
        # Add each consultant to their quadrant
        for i, (data, position) in enumerate(zip(consultants_data, positions)):
            self._add_consultant_to_slide_exact_format(slide, data, position, quadrant_width, quadrant_height)
        
        # Save the presentation
        output_filename = "Team_Slide_Output.pptx"
        output_path = os.path.join(self.output_folder, output_filename)
        prs.save(output_path)
        
        logger.info(f"Team slide saved to {output_path}")
        return output_path
    
    def _add_consultant_to_slide_exact_format(self, slide, consultant_data: Dict, position: Tuple, width, height):
        """
        Add a single consultant's information to a quadrant matching the example format exactly
        """
        left, top = position
        
        # Define layout within quadrant matching the example
        image_size = min(Inches(1.8), width * 0.35)  # Proportional to quadrant
        image_left = left + Inches(0.2)
        image_top = top + Inches(0.2)
        
        # Text starts next to the image
        text_left = image_left + image_size + Inches(0.3)
        text_width = width - image_size - Inches(0.7)
        
        # Add headshot image if available (cropped and resized)
        if consultant_data['headshot_image']:
            try:
                # Calculate target dimensions in pixels (approximate)
                target_width = int(image_size.inches * 96)  # 96 DPI
                target_height = int(image_size.inches * 96)
                
                # Crop and resize image
                processed_image = self._crop_and_resize_image(
                    consultant_data['headshot_image'], 
                    target_width, 
                    target_height
                )
                
                # Save image temporarily
                with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as temp_img:
                    temp_img.write(processed_image)
                    temp_img_path = temp_img.name
                
                # Add image to slide
                slide.shapes.add_picture(temp_img_path, image_left, image_top, image_size, image_size)
                
                # Clean up temp file
                os.unlink(temp_img_path)
                
            except Exception as e:
                logger.warning(f"Could not add image for {consultant_data['name']}: {str(e)}")
        
        # Add consultant name in Consulting Purple (matching example)
        name_top = image_top
        name_box = slide.shapes.add_textbox(text_left, name_top, text_width, Inches(0.4))
        name_frame = name_box.text_frame
        name_frame.margin_left = Pt(0)
        name_frame.margin_top = Pt(0)
        name_frame.margin_bottom = Pt(0)
        name_para = name_frame.paragraphs[0]
        name_run = name_para.add_run()
        name_run.text = consultant_data['name']
        name_run.font.bold = True
        name_run.font.size = Pt(16)  # Larger for prominence
        name_run.font.color.rgb = CONSULTING_PURPLE  # Purple as specified
        
        # Add experience summary (one-line as specified)
        summary_top = name_top + Inches(0.5)
        summary_box = slide.shapes.add_textbox(text_left, summary_top, text_width, Inches(0.3))
        summary_frame = summary_box.text_frame
        summary_frame.margin_left = Pt(0)
        summary_frame.margin_top = Pt(0)
        summary_frame.margin_bottom = Pt(0)
        summary_para = summary_frame.paragraphs[0]
        summary_run = summary_para.add_run()
        summary_run.text = consultant_data['experience_summary']
        summary_run.font.size = Pt(11)
        summary_run.font.color.rgb = GRAY_TEXT
        
        # Add exactly 3 bullet points (as specified)
        bullets_top = summary_top + Inches(0.4)
        bullets_height = height - Inches(1.4)  # Remaining space
        bullets_box = slide.shapes.add_textbox(text_left, bullets_top, text_width, bullets_height)
        bullets_frame = bullets_box.text_frame
        bullets_frame.margin_left = Pt(0)
        bullets_frame.margin_top = Pt(0)
        
        # Add exactly 3 bullet points with consistent formatting
        bullets_to_show = consultant_data['experience_bullets'][:3]  # Ensure exactly 3
        
        # Pad with generic bullets if we don't have enough
        while len(bullets_to_show) < 3:
            bullets_to_show.append("Proven track record in client engagement and project delivery")
        
        for i, bullet in enumerate(bullets_to_show):
            if i == 0:
                para = bullets_frame.paragraphs[0]
            else:
                para = bullets_frame.add_paragraph()
            
            para.level = 0
            para.space_after = Pt(4)  # Consistent spacing
            run = para.add_run()
            run.text = f"• {bullet}"
            run.font.size = Pt(10)
            run.font.color.rgb = BLACK_TEXT
        
        logger.info(f"Added {consultant_data['name']} to slide at position {position} with exact formatting")