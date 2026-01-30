# OCI Metrics to Excel Converter

A web application to convert Oracle Cloud Infrastructure (OCI) metrics CSV files into a formatted Excel spreadsheet with custom instance ordering per project.

## Features

- 📊 Convert OCI metrics CSV files to Excel format
- 🎯 Support for multiple projects (KwikPay, BPL, UIC, Other)
- 🔢 Custom instance ordering per project
- 📈 Handles CPU Mean, CPU Max, Memory Mean, Memory Max metrics
- ✨ Automatic formatting with percentage symbols and center alignment
- 🔄 Smart handling of missing or extra instances
- 💾 Clean, formatted Excel output ready for reporting

## Prerequisites

Before you begin, ensure you have the following installed:
- [Node.js](https://nodejs.org/) (version 16 or higher)
- npm (comes with Node.js)
- Git

## Installation

### 1. Clone the repository
```bash
git clone https://github.com/inura-perera/OCI-Metrics-Updater.git
cd OCI-Metrics-Updater
```

### 2. Install dependencies
```bash
npm install
```

This will install all required packages including:
- React
- Vite
- PapaParse (CSV parser)
- SheetJS (Excel generator)
- Lucide React (icons)

## Running the Application

### Development Mode

To start the development server:
```bash
npm run dev
```

The application will open at `http://localhost:5173/`

### Build for Production

To create a production build:
```bash
npm run build
```

The optimized files will be in the `dist/` folder.

### Preview Production Build

To preview the production build locally:
```bash
npm run preview
```

## How to Use

### Step 1: Select Project
Choose your project from the dropdown menu:
- **KwikPay** - For KwikPay infrastructure metrics
- **BPL** - For BPL infrastructure metrics
- **UIC** - For UIC infrastructure metrics
- **Other** - For projects without predefined ordering

### Step 2: Upload CSV Files
Upload one or more CSV files downloaded from OCI dashboard:
- **CPU Mean** - Average CPU utilization metrics
- **CPU Max** - Maximum CPU utilization metrics
- **Memory Mean** - Average memory utilization metrics
- **Memory Max** - Maximum memory utilization metrics

> **Note:** You can upload any combination of these files. The app will generate Excel with available metrics.

### Step 3: Generate Excel
Click the **"Generate Excel File"** button to create your formatted spreadsheet.

### Step 4: Download
The Excel file will automatically download with the filename format:
```
OCI_Metrics_[ProjectName]_[Date].xlsx
```

## Excel Output Format

The generated Excel file includes:
- Instance names in the first column
- CPU Mean (%), CPU Max (%), Memory Mean (%), Memory Max (%) in subsequent columns
- All percentage values centered and formatted with % symbol
- Bold, centered headers
- Custom instance ordering based on selected project
- Minimum value of 1% for all metrics (values < 1% are rounded up to 1%)

## CSV File Format

Your OCI CSV files should have this structure:
```
group,INSTANCE1,INSTANCE2,INSTANCE3,...
2026-01-26T11:30:00+05:30,98.96,43.46,24.33,...
2026-01-25T11:30:00+05:30,98.36,42.12,23.45,...
```

The app will:
- Automatically detect instance names from column headers
- Extract the latest (most recent) metric values
- Handle missing instances gracefully

## Project Structure
```
OCI-Metrics-Updater/
├── src/
│   ├── App.jsx          # Main application component
│   ├── main.jsx         # Application entry point
│   └── index.css        # Global styles
├── public/              # Static assets
├── index.html           # HTML template
├── package.json         # Dependencies and scripts
├── vite.config.js       # Vite configuration
└── README.md           # This file
```

## Technologies Used

- **React 18** - UI framework
- **Vite** - Build tool and dev server
- **PapaParse** - CSV parsing library
- **SheetJS (xlsx)** - Excel file generation
- **Lucide React** - Icon library
- **Tailwind CSS** - Styling (via CDN)

## Troubleshooting

### Issue: `npm install` fails
**Solution:** Delete `node_modules` folder and `package-lock.json`, then run `npm install` again.
```bash
rm -rf node_modules package-lock.json
npm install
```

### Issue: Port 5173 already in use
**Solution:** Kill the process using that port or specify a different port:
```bash
npm run dev -- --port 3000
```

### Issue: Excel shows "N/A" for some instances
**Solution:** This means the instance exists in your predefined list but wasn't found in the uploaded CSV files. This is expected behavior for instances that are not currently active.

### Issue: Folder name with special characters (like R&D) causes errors
**Solution:** Avoid using special characters (`&`, `%`, `!`, etc.) in folder paths. Use alphanumeric names or underscores/hyphens instead.

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License.

## Support

For issues, questions, or suggestions, please open an issue on GitHub.

## Acknowledgments

- Oracle Cloud Infrastructure for the metrics data
- The React and Vite communities for excellent tools and documentation
