# CSV Classifier Application

This project is a Flask web application that allows users to upload CSV files, processes the data to classify it, and returns the results as an Excel file. The application features an intuitive user interface with a logo and a background image.

## Features

- Upload CSV files for classification.
- Process the uploaded data using predefined classification rules.
- Download the classified data as an Excel file.
- User-friendly interface with visual enhancements.

## Project Structure

```
csv-classifier-app
├── app.py                # Main entry point of the Flask application
├── static                # Static files (CSS, JS, images)
│   ├── css
│   │   └── style.css     # Styles for the application
│   ├── js
│   │   └── main.js       # JavaScript for client-side interactivity
│   └── images
│       ├── logo.png      # Logo displayed at the top of the application
│       └── background.jpg # Background image for the application
├── templates             # HTML templates for rendering pages
│   ├── base.html         # Base template with common structure
│   ├── index.html        # Main page for file uploads
│   └── results.html      # Page to display results and download link
├── uploads               # Directory for uploaded files
│   └── .gitkeep          # Keeps the uploads directory in version control
├── processed             # Directory for processed files
│   └── .gitkeep          # Keeps the processed directory in version control
├── utils                 # Utility functions and classifiers
│   ├── __init__.py       # Marks the utils directory as a Python package
│   └── classifier.py     # Contains classification logic
├── requirements.txt      # Lists project dependencies
├── config.py             # Configuration settings for the Flask app
└── README.md             # Project documentation
```

## Installation

1. Clone the repository:
   ```
   git clone <repository-url>
   cd csv-classifier-app
   ```

2. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

3. Run the application:
   ```
   python app.py
   ```

4. Open your web browser and go to `http://127.0.0.1:5000` to access the application.

## Usage

- On the main page, upload your CSV file using the provided form.
- After processing, you will be redirected to a results page where you can download the classified Excel file.

## Contributing

Contributions are welcome! Please feel free to submit a pull request or open an issue for any suggestions or improvements.

## License

This project is licensed under the MIT License. See the LICENSE file for details.