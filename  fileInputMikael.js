document.getElementById('processButtonMikael').addEventListener('click', () => {
    const fileInput = document.getElementById('fileInputMikael');
    
    if (!fileInput.files.length) {
        alert('Please select a file first.');
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = (event) => {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const sourceSheet = workbook.Sheets[workbook.SheetNames[0]];
            const sourceData = XLSX.utils.sheet_to_json(sourceSheet);

            const formattedData = sourceData.map(item => {
                const bodyFabric = item['Body/Fabric'] || '';
                const handleTitle = bodyFabric.replace(/\s+/g, '').toLowerCase();
                const tags = `${bodyFabric.replace(/\s+/g, '')},Mikael,${item['Description'] || ''}`;

                return {
                    Handle: handleTitle,
                    Command: 'MERGE',
                    Title: handleTitle,
                    'Body HTML': item['Material Composition by %'] || '',
                    Vendor: item['Division Name'] || '',
                    Type: item['Category'] || '',
                    Tags: tags,
                    'Tags Command': 'REPLACE',
                    Status: 'active',
                    'Total Inventory Qty': item['Booked Units'] || '0',
                    'Image Src': '',
                    'Image Command': 'MERGE',
                    'Option1 Name': 'Color',
                    'Option1 Value': item['Color Description'] || '',
                    'Option2 Name': 'Size',
                    'Option2 Value': item['Size'] || '',
                    'Variant SKU': item['UPC'] || '',
                    'Variant Barcode': item['UPC'] || '',
                    'Variant Weight': '2',
                    'Variant Weight Unit': 'lb',
                    'Variant Price': item['Cost'] || '',
                    'Variant Compare At Price': item['Cost'] || '',
                    'Variant Taxable': 'TRUE',
                    'Variant Inventory Tracker': 'shopify',
                    'Variant Inventory Policy': 'deny',
                    'Variant Fulfillment Service': 'manual',
                    'Variant Inventory Qty': item['Booked Units'] || '0',
                    'Variant Country of Origin': item['Country of Origin'] || '',
                    'Variant Metafield: mm-google-shopping.age_group [single_line_text_field]': 'Adult',
                    'Variant Metafield: mm-google-shopping.condition [single_line_text_field]': 'New',
                    'Variant Metafield: mm-google-shopping.gender [single_line_text_field]': 'Female',
                    'Variant Metafield: mm-google-shopping.custom_label_1 [single_line_text_field]': item['Size'] || '',
                    'Variant Metafield: mm-google-shopping.custom_label_2 [single_line_text_field]': item['Color Description'] || ''
                };
            });

            const csv = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(formattedData));

            const blob = new Blob([csv], { type: 'text/csv' });

            const downloadLink = document.getElementById('downloadLinkMikael');
            downloadLink.href = URL.createObjectURL(blob);
            downloadLink.download = 'formatted_inventory.csv';
            downloadLink.style.display = 'block';
            downloadLink.innerHTML = 'Download Formatted CSV File';
        } catch (error) {
            alert('An error occurred while processing the file.');
            console.error(error);
        }
    };

    reader.readAsArrayBuffer(file);
});

document.getElementById('downloadLinkMikael').style.display = 'none';
