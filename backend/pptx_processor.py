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
import shutil

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
            cv_files = [f for f in os.listdir(self.cvs_folder) if f.endswith('.pptx') and not f.startswith('CV_Placeholder')]
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
        
    def extract_consultant_data_from_template(self, cv_filepath: str, consultant_name: str) -> Dict:
        """
        Extract consultant data from a CV PowerPoint file using the CV_Placeholder structure as reference
        Returns: {
            'name': str,
            'role': str, 
            'location': str,
            'experience_bullets': List[str],
            'headshot_image': bytes or None
        }
        """
        logger.info(f"Extracting data from {cv_filepath} using template structure")
        
        try:
            prs = Presentation(cv_filepath)
            
            # Assume we're working with the first slide
            if len(prs.slides) == 0:
                raise ValueError(f"No slides found in {cv_filepath}")
                
            slide = prs.slides[0]
            
            # Initialize data
            consultant_data = {
                'name': consultant_name,
                'role': "Senior Consultant",
                'location': "Global",
                'experience_bullets': [],
                'headshot_image': None
            }
            
            # Extract data from shapes based on CV_Placeholder structure
            for i, shape in enumerate(slide.shapes):
                try:
                    # Shape 7: IMAGE (headshot) - based on our template analysis
                    if hasattr(shape, 'image') and consultant_data['headshot_image'] is None:
                        image_stream = io.BytesIO(shape.image.blob)
                        consultant_data['headshot_image'] = image_stream.getvalue()
                        logger.info(f"Extracted headshot image from shape {i}")
                    
                    # Shape 4: Name, position, location - based on our template analysis  
                    elif hasattr(shape, 'text') and shape.text.strip():
                        text = shape.text.strip()
                        
                        # Check if this contains name/position info (usually has "Position" or office locations)
                        if any(keyword in text.lower() for keyword in ['position', 'office', 'location', 'germany', 'london', 'new york', 'paris', 'berlin', 'zurich', 'geneva', 'munich']):
                            lines = [line.strip() for line in text.split('\n') if line.strip()]
                            if lines:
                                # Find the name line - usually contains comma and proper name structure
                                for line in lines:
                                    if ',' in line and any(char.isalpha() for char in line):
                                        name_parts = line.split(',')
                                        if len(name_parts) >= 2:
                                            # Format: "Last Name, First Name" or similar
                                            last_name = name_parts[0].strip()
                                            first_name = name_parts[1].strip()
                                            # Only update if this looks like a proper name (not random text)
                                            if len(last_name) < 50 and len(first_name) < 50 and not any(keyword in line.lower() for keyword in ['university', 'msc', 'ba', 'phd', 'degree']):
                                                consultant_data['name'] = f"{first_name} {last_name}"
                                                break
                                
                                # Look for position and location in all lines
                                for line in lines:
                                    if any(keyword in line.lower() for keyword in ['consultant', 'manager', 'director', 'analyst', 'partner']) and len(line) < 100:
                                        if consultant_data['role'] == "Senior Consultant":  # Only update if still default
                                            consultant_data['role'] = line.strip()
                                    elif any(keyword in line.lower() for keyword in ['germany', 'london', 'new york', 'paris', 'berlin', 'zurich', 'geneva', 'munich']) and len(line) < 100:
                                        if consultant_data['location'] == "Global":  # Only update if still default
                                            consultant_data['location'] = line.strip()
                        
                        # Shape 1: "Selected consulting engagement experience" - extract bullet points
                        elif 'consulting engagement experience' in text.lower() or 'consulting experience' in text.lower():
                            lines = [line.strip() for line in text.split('\n') if line.strip()]
                            bullets = []
                            
                            for line in lines:
                                # Skip header lines
                                if 'consulting engagement experience' in line.lower() or 'take 3 bullet' in line.lower():
                                    continue
                                    
                                # Look for actual bullet points (meaningful content lines)
                                if len(line) > 20 and not line.startswith('Take '):  # Avoid instruction text
                                    # Clean up bullet formatting
                                    clean_line = line.lstrip('•-▪◦→ ').strip()
                                    if clean_line and len(clean_line) > 20:
                                        bullets.append(clean_line)
                            
                            # Take only first 3 bullets as required
                            consultant_data['experience_bullets'] = bullets[:3]
                            logger.info(f"Extracted {len(consultant_data['experience_bullets'])} experience bullets")
                        
                except Exception as e:
                    logger.warning(f"Error processing shape {i}: {str(e)}")
                    continue
            
            # Ensure we have exactly 3 bullet points
            while len(consultant_data['experience_bullets']) < 3:
                consultant_data['experience_bullets'].append("Proven track record in client engagement and project delivery")
            
            consultant_data['experience_bullets'] = consultant_data['experience_bullets'][:3]
            
            logger.info(f"Extracted data - Name: {consultant_data['name']}, Role: {consultant_data['role']}, "
                       f"Location: {consultant_data['location']}, Bullets: {len(consultant_data['experience_bullets'])}")
            
            return consultant_data
            
        except Exception as e:
            logger.error(f"Error extracting data from {cv_filepath}: {str(e)}")
            raise
    
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

    def create_team_slide(self, names: List[str]) -> str:
        """
        Create a team slide using template files instead of building from scratch.
        Uses CV_Placeholder.pptx as reference for parsing consultant CVs
        and Output_Example_Placeholder_Logic.pptx as the base template.
        
        Args:
            names: List of consultant names to include in the team slide
            
        Returns:
            Path to the generated Team_Slide_Output.pptx file
        """
        logger.info(f"Creating team slide for consultants: {names}")
        
        # Find and extract data from consultant CV files
        consultants_data = []
        for name in names:
            filename = self.find_cv_file(name)
            if not filename:
                logger.warning(f"CV file not found for {name}, using placeholder data")
                # Create placeholder data for missing consultant
                consultants_data.append({
                    'name': name,
                    'role': "Senior Consultant",
                    'location': "Global",
                    'experience_bullets': [
                        "Extensive experience in strategic consulting",
                        "Proven track record in client engagement",
                        "Specialized in project delivery and transformation"
                    ],
                    'headshot_image': None
                })
            else:
                cv_filepath = os.path.join(self.cvs_folder, filename)
                try:
                    data = self.extract_consultant_data_from_template(cv_filepath, name)
                    consultants_data.append(data)
                except Exception as e:
                    logger.error(f"Failed to extract data from {filename}: {str(e)}")
                    # Use placeholder data if extraction fails
                    consultants_data.append({
                        'name': name,
                        'role': "Senior Consultant", 
                        'location': "Global",
                        'experience_bullets': [
                            "Extensive experience in strategic consulting",
                            "Proven track record in client engagement", 
                            "Specialized in project delivery and transformation"
                        ],
                        'headshot_image': None
                    })
        
        # Load the output template
        template_path = os.path.join(self.examples_folder, 'Outpout_Example_Placeholder_Logic.pptx')
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Output template not found: {template_path}")
        
        logger.info(f"Loading output template from {template_path}")
        prs = Presentation(template_path)
        slide = prs.slides[0]
        
        # Map consultant data to template placeholders
        # Based on our analysis: shapes 8, 11, 12, 13 contain consultant text
        # shapes 10, 14, 15, 16 contain consultant images
        consultant_text_shapes = [8, 11, 12, 13]
        consultant_image_shapes = [10, 14, 15, 16]
        
        for i, consultant_data in enumerate(consultants_data[:4]):  # Limit to 4 consultants
            try:
                # Update text placeholder
                if i < len(consultant_text_shapes):
                    text_shape_idx = consultant_text_shapes[i]
                    if text_shape_idx < len(slide.shapes):
                        text_shape = slide.shapes[text_shape_idx]
                        if hasattr(text_shape, 'text'):
                            # Create consultant text content
                            consultant_text = f"{consultant_data['name']}\n{consultant_data['role']}, {consultant_data['location']}\n"
                            consultant_text += "x+ years of consulting experience\n\n"
                            for bullet in consultant_data['experience_bullets']:
                                consultant_text += f"• {bullet}\n"
                            
                            text_shape.text = consultant_text
                            logger.info(f"Updated text for consultant {i+1}: {consultant_data['name']}")
                
                # Update image placeholder
                if i < len(consultant_image_shapes) and consultant_data['headshot_image']:
                    image_shape_idx = consultant_image_shapes[i]
                    if image_shape_idx < len(slide.shapes):
                        try:
                            # Get the current image shape to determine size and position
                            current_shape = slide.shapes[image_shape_idx]
                            
                            # Get position and size
                            left = current_shape.left
                            top = current_shape.top
                            width = current_shape.width
                            height = current_shape.height
                            
                            # Process the headshot image
                            target_width = int(width.inches * 96)  # Convert to pixels (96 DPI)
                            target_height = int(height.inches * 96)
                            
                            processed_image = self._crop_and_resize_image(
                                consultant_data['headshot_image'],
                                target_width,
                                target_height
                            )
                            
                            # Save processed image temporarily
                            with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as temp_img:
                                temp_img.write(processed_image)
                                temp_img_path = temp_img.name
                            
                            # Remove old image shape
                            slide.shapes._spTree.remove(current_shape._element)
                            
                            # Add new image
                            slide.shapes.add_picture(temp_img_path, left, top, width, height)
                            
                            # Clean up temp file
                            os.unlink(temp_img_path)
                            
                            logger.info(f"Updated image for consultant {i+1}: {consultant_data['name']}")
                            
                        except Exception as e:
                            logger.warning(f"Failed to update image for consultant {i+1}: {str(e)}")
                
            except Exception as e:
                logger.error(f"Failed to update consultant {i+1} data: {str(e)}")
                continue
        
        # Save the final presentation
        output_filename = "Team_Slide_Output.pptx"
        output_path = os.path.join(self.output_folder, output_filename)
        prs.save(output_path)
        
        logger.info(f"Team slide saved to {output_path}")
        return output_path

    # Keep the old method name for backward compatibility
    def generate_team_slide(self, consultant_names: List[str], filenames: List[str] = None) -> str:
        """
        Backward compatibility wrapper for create_team_slide
        """
        return self.create_team_slide(consultant_names)