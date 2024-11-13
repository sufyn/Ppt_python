# PowerPoint Text Customization

This project demonstrates how to automate the creation of a PowerPoint presentation using Python, where the content for each slide is sourced from external text files and custom font styles are applied to the text.

## Overview

The project uses the `python-pptx` library to create a PowerPoint presentation, adding slides and populating them with content from text files. The script also applies a custom font to the text content of each slide.

## Features

- Create PowerPoint slides dynamically from input text files.
- Apply a custom font style to the text content.
- Customizable slide layout and text box size.
- Support for multiple slides with different text content and styles.

## Requirements

- Python 3.x
- `python-pptx` library
- A custom font file (e.g., `Love Ya Like A Sister.ttf`)
- Sample input text files (e.g., `sample_slide1_input.txt`, `sample_slide2_input.txt`)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/pptx_text_customization.git
   cd pptx_text_customization
   ```

2. Install the necessary Python package:
   ```bash
   pip install python-pptx
   ```

3. Ensure you have the necessary input text files (`sample_slide1_input.txt`, `sample_slide2_input.txt`) and font file (`sample_font_file.ttf`) in the same directory as the script.

## How It Works

1. **Slide Creation:** The script creates a presentation and adds two slides.
2. **Reading Content:** For each slide, the content is read from external text files (`sample_slide1_input.txt` and `sample_slide2_input.txt`).
3. **Text Box Creation:** Text boxes are added to the slides, and the content from the text files is inserted into them.
4. **Font Customization:** The script applies a custom font (`Love Ya Like A Sister`) to the text content of each slide.
5. **Saving the Presentation:** The PowerPoint file is saved as `test.pptx`.

## Code

```python
import collections.abc
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR

# Define the path to the input text files and font file
slide1_input_file = 'sample_slide1_input.txt'
slide2_input_file = 'sample_slide2_input.txt'
font_file = 'sample_font_file.ttf'


prs = Presentation()

blank_slide_layout = prs.slide_layouts[6]
slide1 = prs.slides.add_slide(blank_slide_layout)

# Read content from sample_slide1_input.txt
with open(slide1_input_file, 'r') as file:
    slide1_content = file.read()

width = Inches(10)
left = top =  right = Inches(0.1)
txBox = slide1.shapes.add_textbox(left, top, width, right)
tf = txBox.text_frame

tf.text = slide1_content
tf.word_wrap = True

# Set the font style for Slide 1
for paragraph in tf.paragraphs:
    for run in paragraph.runs:
        run.font.file = font_file
        run.font.name = 'Love Ya Like A Sister'
        run.font.size = Pt(13)

# Second Slide
blank_slide_layout = prs.slide_layouts[6]
slide2 = prs.slides.add_slide(blank_slide_layout)

# Read content from sample_slide2_input.txt
with open(slide2_input_file, 'r') as file:
    slide2_content = file.read()

width = Inches(10)
left = top =  right  = Inches(0.1)
txBox = slide2.shapes.add_textbox(left, top, width, right)
tf = txBox.text_frame

tf.text = slide2_content
tf.word_wrap = True

# Set the font style for Slide 2
for paragraph in tf.paragraphs:
    for run in paragraph.runs:
        run.font.file = font_file
        run.font.name = 'Love Ya Like A Sister'
        run.font.size = Pt(23)

# Save the presentation
prs.save('test.pptx')
```

## How to Run

1. Place the necessary input text files (`sample_slide1_input.txt`, `sample_slide2_input.txt`) and the font file (`sample_font_file.ttf`) in the project directory.
2. Run the script:
   ```bash
   python pptx_text_customization.py
   ```
3. The script will generate a PowerPoint file named `test.pptx`.

## Example of Output

After running the script, a PowerPoint presentation will be generated with the following characteristics:
- **Slide 1**: The content from `sample_slide1_input.txt` will be displayed with the custom font style (`Love Ya Like A Sister` at 13pt size).
- **Slide 2**: The content from `sample_slide2_input.txt` will be displayed with the custom font style (`Love Ya Like A Sister` at 23pt size).

## Screenshot

Below is a screenshot showing a sample output from the script:

![Sample Output](https://github.com/yourusername/pptx_text_customization/assets/yourimage)

## Contributing

Feel free to fork this repository, report issues, or submit pull requests to improve the functionality.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
