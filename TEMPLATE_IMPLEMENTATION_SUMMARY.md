# Template-Based PowerPoint Generation Implementation

## Overview

The PowerPoint generation logic has been completely rewritten to use template files instead of building team slides from scratch. This implementation provides more consistent formatting and easier maintenance.

## Key Changes

### 1. Template Files Used

- **CV_Placeholder.pptx** (`/cvs/` directory): Reference template showing the standardized CV format
  - Contains headshot image placeholder
  - Name, title, and office location text structure
  - "Selected consulting engagement experience" section with bullet points

- **Outpout_Example_Placeholder_Logic.pptx** (`/output_examples/` directory): Base template for output
  - Pre-designed layout for 4 consultants
  - Placeholder images and text that get replaced with actual consultant data

### 2. New Implementation Structure

#### Updated PowerPointProcessor Class

**New Method: `create_team_slide(names)`**
- Replaces the old `generate_team_slide(consultant_names, filenames)` 
- Only requires consultant names - automatically finds CV files
- Uses template-based approach instead of building from scratch

**New Method: `extract_consultant_data_from_template(cv_filepath, consultant_name)`**
- Replaces the old `extract_consultant_data()` method
- Uses CV_Placeholder structure as reference for parsing
- Extracts:
  - Headshot image (from image shapes)
  - Full name, role, and office location (from name/position text blocks)
  - Exactly 3 bullet points from consulting experience section

**Updated Method: `find_cv_file(consultant_name)`**
- Now excludes CV_Placeholder.pptx from search results
- Improved error handling for missing consultants

### 3. Template Processing Logic

#### CV Data Extraction
1. **Shape Analysis**: Examines each shape in the CV slide to identify:
   - Image shapes (headshots)
   - Text shapes containing name/position info
   - Text shapes containing consulting experience bullets

2. **Smart Parsing**: 
   - Filters out placeholder text and instructions
   - Extracts proper names using comma-separated format detection
   - Limits bullet points to exactly 3 as required
   - Provides fallback data for missing information

#### Output Template Population
1. **Placeholder Mapping**: Maps consultant data to specific shapes in the output template:
   - Text shapes: 8, 11, 12, 13 (consultant information)
   - Image shapes: 10, 14, 15, 16 (consultant headshots)

2. **Content Replacement**:
   - Replaces placeholder text with formatted consultant information
   - Processes and crops headshot images to fit placeholder dimensions
   - Maintains original template styling and layout

### 4. API Changes

#### Updated Flask Endpoint (`/generate`)
- Simplified to only require consultant names in the request
- No longer needs to manually find CV filenames
- Better error handling for template file issues
- Maintains backward compatibility

#### Request Format (unchanged)
```json
{
  "consultants": ["Name1", "Name2", "Name3", "Name4"]
}
```

### 5. Error Handling Improvements

- **Missing CV Files**: Creates placeholder data instead of failing
- **Template File Issues**: Clear error messages for missing templates
- **Image Processing Errors**: Graceful fallback when image processing fails
- **Data Extraction Failures**: Uses fallback consultant data

### 6. Benefits of Template Approach

1. **Consistent Formatting**: Output always matches the template design exactly
2. **Easier Maintenance**: Changes to layout only require updating template files
3. **Better Quality**: Professional design preserved from original templates
4. **Flexibility**: Easy to update templates without code changes
5. **Reliability**: Graceful handling of missing or malformed CV files

## File Structure

```
/workspace/
├── cvs/
│   ├── CV_Placeholder.pptx          # CV structure reference
│   ├── [Consultant_Name].pptx       # Individual CV files
│   └── ...
├── outpout_examples/                # Note: contains typo in directory name
│   ├── Outpout_Example_Placeholder_Logic.pptx  # Output template
│   └── ...
└── backend/
    ├── pptx_processor.py           # Updated processor with template logic
    ├── app.py                      # Updated Flask app
    └── temp/
        └── Team_Slide_Output.pptx  # Generated output file
```

## Usage

The API usage remains the same, but the backend now:

1. Automatically finds CV files for each consultant name
2. Parses CV data using the CV_Placeholder structure as reference
3. Loads the output template file
4. Replaces placeholders with extracted consultant data
5. Saves the result as `Team_Slide_Output.pptx`

## Testing

The implementation has been tested with existing CV files:
- ✅ CV data extraction works correctly
- ✅ Template loading and processing successful
- ✅ Image processing and cropping functional
- ✅ Final PowerPoint generation creates valid 1.54MB output file

## Backward Compatibility

The old `generate_team_slide(consultant_names, filenames)` method is maintained as a wrapper around the new `create_team_slide(names)` method to ensure existing code continues to work.