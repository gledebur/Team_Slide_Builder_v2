from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import logging
from pptx_processor import PowerPointProcessor

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# Configuration
CVS_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'cvs')
OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp')
OUTPUT_EXAMPLES_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'outpout_examples')

# Ensure output folder exists
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({"status": "healthy"}), 200

@app.route('/generate-slide', methods=['POST'])
def generate_slide():
    """
    Generate team slide endpoint for V0 frontend.
    When specific consultant names are provided, returns pre-generated PowerPoint file.
    Expected JSON payload: {"consultants": ["Name1", "Name2", "Name3", "Name4"]}
    """
    try:
        # Get consultant names from request
        data = request.get_json()
        if not data or 'consultants' not in data:
            return jsonify({"error": "Missing 'consultants' field in request body"}), 400
        
        consultant_names = data['consultants']
        
        # Validate we have exactly 4 consultants
        if len(consultant_names) != 4:
            return jsonify({"error": "Exactly 4 consultant names are required"}), 400
        
        # Validate all names are non-empty
        if not all(name.strip() for name in consultant_names):
            return jsonify({"error": "All consultant names must be non-empty"}), 400
        
        logger.info(f"Processing consultants for slide generation: {consultant_names}")
        
        # Check if the provided names match the specific ones for pre-generated file
        expected_names = [
            "Caledonia Trapp",
            "Benjamin Reinitzer", 
            "Benedict Wolske",
            "Gregor Ledebur"
        ]
        
        # Normalize names for comparison (strip whitespace and compare case-insensitively)
        normalized_input = [name.strip() for name in consultant_names]
        normalized_expected = [name.strip() for name in expected_names]
        
        # Check if the input names match the expected ones (order doesn't matter)
        if set(normalized_input) == set(normalized_expected):
            logger.info("Consultant names match expected list, returning pre-generated PowerPoint file")
            
            # Path to the pre-generated PowerPoint file
            pregenerated_file = os.path.join(OUTPUT_EXAMPLES_FOLDER, 'Outpout_Example.pptx')
            
            # Check if the file exists
            if not os.path.exists(pregenerated_file):
                logger.error(f"Pre-generated file not found: {pregenerated_file}")
                return jsonify({"error": "Pre-generated PowerPoint file not found"}), 404
            
            # Return the pre-generated file
            return send_file(
                pregenerated_file,
                as_attachment=True,
                download_name='Team_Slide_Output.pptx',
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )
        else:
            # For any other names, return an error or could implement actual CV processing
            logger.info("Consultant names don't match expected list")
            return jsonify({
                "error": "CV processing not implemented for these consultants. Please use the specific consultant names from the demo."
            }), 400
        
    except Exception as e:
        logger.error(f"Error in generate-slide endpoint: {str(e)}")
        return jsonify({"error": f"Failed to generate team slide: {str(e)}"}), 500

@app.route('/generate', methods=['POST'])
def generate_team_slide():
    """
    Generate team slide from consultant names using template files
    Expected JSON payload: {"consultants": ["Name1", "Name2", "Name3", "Name4"]}
    """
    try:
        # Get consultant names from request
        data = request.get_json()
        if not data or 'consultants' not in data:
            return jsonify({"error": "Missing 'consultants' field in request body"}), 400
        
        consultant_names = data['consultants']
        
        # Validate we have exactly 4 consultants
        if len(consultant_names) != 4:
            return jsonify({"error": "Exactly 4 consultant names are required"}), 400
        
        # Validate all names are non-empty
        if not all(name.strip() for name in consultant_names):
            return jsonify({"error": "All consultant names must be non-empty"}), 400
        
        logger.info(f"Processing consultants: {consultant_names}")
        
        # Initialize PowerPoint processor
        processor = PowerPointProcessor(CVS_FOLDER, OUTPUT_FOLDER, OUTPUT_EXAMPLES_FOLDER)
        
        # Generate team slide using template files - no need to find filenames manually
        output_file = processor.create_team_slide(consultant_names)
        
        # Return the generated file
        return send_file(
            output_file,
            as_attachment=True,
            download_name='Team_Slide_Output.pptx',
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
    except FileNotFoundError as e:
        logger.error(f"Template file not found: {str(e)}")
        return jsonify({"error": f"Template file not found: {str(e)}"}), 404
        
    except Exception as e:
        logger.error(f"Error generating team slide: {str(e)}")
        return jsonify({"error": f"Failed to generate team slide: {str(e)}"}), 500

@app.route('/list-cvs', methods=['GET'])
def list_cvs():
    """List available CV files for debugging"""
    try:
        cv_files = [f for f in os.listdir(CVS_FOLDER) if f.endswith('.pptx') and not f.startswith('CV_Placeholder')]
        return jsonify({"cv_files": cv_files}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)