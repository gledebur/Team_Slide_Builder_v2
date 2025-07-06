# Team Slide Generator - Complete Setup Guide

This system automatically generates professional team slides from consultant CV PowerPoint files.

## Architecture

- **Frontend**: Next.js application (already built with v0.dev)
- **Backend**: Flask API that processes PowerPoint files
- **File Structure**:
  - `cvs/` - Contains individual consultant CV slides (.pptx files)
  - `outpout_examples/` - Contains example output format
  - `backend/` - Flask API server

## Setup Instructions

### 1. Backend Setup

```bash
# Navigate to backend directory
cd backend

# Option 1: Use the automated script
./run.sh

# Option 2: Manual setup
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python app.py
```

The Flask server will start on `http://localhost:5000`

### 2. Frontend Setup

```bash
# In the root directory
npm install
npm run dev
```

The Next.js app will start on `http://localhost:3000`

## How It Works

### 1. User Input
Users enter 4 consultant names in the web interface:
- Gregor Ledebur
- Benedict Wolske  
- Benjamin Reinitzer
- Caledonia Trapp

### 2. File Matching
The system uses flexible matching to find CV files:
- **Standard rule**: `"Name Surname"` → `"Name_Surname.pptx"`
- **Fuzzy matching**: If exact match fails, finds files containing all name parts
- **Example**: "Gregor Ledebur" matches "Gregor_Ledebur-Wicheln.pptx"

### 3. Data Extraction
From each CV slide, extracts:
- **Headshot**: First image shape in the slide
- **Name, Role, Location**: Parsed from text boxes
- **Experience bullets**: 3-4 key bullet points

### 4. Team Slide Generation
Creates a new PowerPoint with 2x2 layout:
- Each quadrant contains one consultant
- Left side: headshot image
- Right side: name (bold), role/location (italic), experience bullets

### 5. File Download
Returns the generated `Team_Slide_Output.pptx` file for download

## API Endpoints

### `POST /generate`
Generates team slide from consultant names.

**Request:**
```json
{
  "consultants": ["Name1", "Name2", "Name3", "Name4"]
}
```

**Response:** PowerPoint file download

### `GET /health`
Health check endpoint.

### `GET /list-cvs`
Lists available CV files (for debugging).

## File Requirements

### CV Files (in `cvs/` folder)
- Must be `.pptx` format
- First slide should contain:
  - At least one image (consultant headshot)
  - Text boxes with name, role, location
  - Bullet points for experience

### Naming Convention
CV files should follow the pattern: `Firstname_Lastname.pptx`
- Spaces replaced with underscores
- Hyphens preserved (e.g., `Gregor_Ledebur-Wicheln.pptx`)

## Troubleshooting

### Common Issues

1. **CV file not found**
   - Check filename matches consultant name
   - Use `/list-cvs` endpoint to see available files
   - Verify files are in the `cvs/` folder

2. **Image extraction fails**
   - Ensure CV slides contain image shapes
   - Check image format compatibility

3. **Text parsing issues**
   - Verify CV slides have text boxes with consultant info
   - Check for special characters in text

4. **CORS errors**
   - Ensure Flask server is running on port 5000
   - Frontend should be on port 3000

### Debug Mode
The Flask app runs in debug mode by default. Check console logs for detailed error information.

## Dependencies

### Backend (Python)
- Flask 3.1.0 - Web framework
- python-pptx 1.0.2 - PowerPoint manipulation
- Pillow 10.4.0 - Image processing
- flask-cors 5.0.0 - CORS handling

### Frontend (Node.js)
- Next.js 15.2.4 - React framework
- Various UI components (@radix-ui/*)
- Tailwind CSS for styling

## Testing

Test the system with the provided consultant names:
1. Start both backend and frontend servers
2. Enter the 4 consultant names in the web interface
3. Click "Generate Team Slide"
4. Download and verify the generated PowerPoint file

The output should match the format shown in `outpout_examples/Outpout_Example.pptx`.