# CV Usage Guide

This guide explains how to use and update the CV documents in this repository.

## Available Formats

### Markdown CV (`cv.md`)

The Markdown version is ideal for:
- Viewing directly on GitHub
- Quick edits and updates
- Version control and tracking changes

### Word CV (`Raja_Ahsan_Ali_CV.docx`)

The Word document is ideal for:
- Printing a professional copy
- Sending to employers via email
- Uploading to job portals

## How to Update the CV

### Method 1: Edit Markdown Directly

1. Open `docs/cv.md`
2. Make your changes
3. Commit and push

### Method 2: Regenerate Word Document

If you make changes to the CV content, you can regenerate the Word document:

```bash
# Install dependencies
pip install python-docx

# Run the script
python create_cv.py
```

The script will create/update `docs/Raja_Ahsan_Ali_CV.docx`.

## CV Sections

The CV includes the following sections:

1. **Personal Information** - Name, age, height, weight, QID, location
2. **Contact Information** - Email and phone number
3. **Professional Summary** - Brief overview of experience and goals
4. **Work Experience** - Previous positions and responsibilities
5. **Education** - Academic qualifications
6. **Skills** - Key competencies and abilities
7. **Languages** - Language proficiency levels
8. **References** - Available upon request

## Profile Photo

The profile photo is stored at `docs/images/profile_photo.jpeg`. To update:

1. Replace the file with a new photo (keep the same filename)
2. Recommended size: 150x150 pixels or similar aspect ratio
3. Use a professional headshot

## Tips for CV Updates

- Keep work experience descriptions concise and action-oriented
- Use bullet points for easy readability
- Update contact information promptly if it changes
- Add new skills as you acquire them
- Maintain consistent formatting throughout
