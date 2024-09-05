const pdfParse = require('pdf-parse');
const ExcelJS = require('exceljs');
const fs = require('fs');

// 1. Parse the PDF
const parsePdf = (pdfBuffer) => {
    return pdfParse(pdfBuffer).then(data => {
        // Process the text from the PDF to extract the table data
        const pdfText = data.text;
        const rows = pdfText.split('\n').filter(line => line.match(/^\d+/)); // Get rows starting with beer details
        const beers = rows.map(row => {
            const [name, style, sud, plech, sklo75, sklo33] = row.split(/\s+/);
            return { name, style, sud, plech, sklo75, sklo33 };
        });
        return beers;
    });
};

// 2. Create Excel File
const createExcelFile = async (beerList) => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Beer Order');

    // Create headers
    sheet.columns = [
        { header: 'Název', key: 'name', width: 20 },
        { header: 'Styl', key: 'style', width: 20 },
        { header: 'Sud cena/0,5l', key: 'sud', width: 15 },
        { header: 'Plech 0,5l', key: 'plech', width: 15 },
        { header: 'Sklo 0,75l', key: 'sklo75', width: 15 },
        { header: 'Sklo 0,33l', key: 'sklo33', width: 15 },
        { header: 'Order for Person 1 (sud)', key: 'person1_sud', width: 20 },
        { header: 'Order for Person 1 (plech)', key: 'person1_plech', width: 20 },
        { header: 'Order for Person 1 (sklo75)', key: 'person1_sklo75', width: 20 },
        { header: 'Order for Person 1 (sklo33)', key: 'person1_sklo33', width: 20 },
        { header: 'Total Price for Person 1', key: 'total_person1', width: 20 }
    ];

    // 3. Add beers to the Excel file
    beerList.forEach(beer => {
        sheet.addRow({
            name: beer.name,
            style: beer.style,
            sud: beer.sud === 'X' ? '-' : beer.sud,
            plech: beer.plech === 'X' ? '-' : beer.plech,
            sklo75: beer.sklo75 === 'X' ? '-' : beer.sklo75,
            sklo33: beer.sklo33 === 'X' ? '-' : beer.sklo33,
            total_person1: `=(${beer.sud}*G{row})+(${beer.plech}*H{row})+(${beer.sklo75}*I{row})+(${beer.sklo33}*J{row})`
        });
    });

    // 4. Save the file
    await workbook.xlsx.writeFile('BeerOrder.xlsx');
};

// Main Execution
const pdfPath = '/mnt/data/Obora VOC    2. - 8. 9. 2024.pdf';
const pdfBuffer = fs.readFileSync(pdfPath);

parsePdf(pdfBuffer).then(beerList => {
    createExcelFile(beerList);
}).catch(err => console.error(err));