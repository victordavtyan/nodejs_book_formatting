const express = require('express');
const multer = require('multer');
const AdmZip = require('adm-zip');
const fs = require('fs');
const path = require('path');
const xml2js = require('xml2js');

const app = express();
const upload = multer({ dest: 'uploads/' });

const PORT = 3000;

app.use(express.static('public')); // Serve static files from "public" folder
app.use(express.urlencoded({ extended: true })); // To parse form data

async function replaceBackgroundImage(odtPath, newImagePath, outputPath) {
    const zip = new AdmZip(odtPath);
    const tempDir = `temp_${Date.now()}`;

    // Extract ODT contents
    zip.extractAllTo(tempDir, true);

    const picturesDir = path.join(tempDir, 'Pictures');
    const stylesXmlPath = path.join(tempDir, 'styles.xml');

    // Check if a Pictures folder exists
    if (!fs.existsSync(picturesDir)) {
        throw new Error('No Pictures folder found in the uploaded ODT.');
    }

    // Replace the first image in the Pictures folder with the new image
    const pictureFiles = fs.readdirSync(picturesDir);
    const imageFile = pictureFiles.find(file => file.match(/\.(png|jpg|jpeg)$/i));

    if (!imageFile) {
        throw new Error('No image found in the Pictures folder.');
    }

    // Replace the image
    const newImageName = path.basename(imageFile); // Use the same name as the original image
    fs.copyFileSync(newImagePath, path.join(picturesDir, newImageName));

    // Update `styles.xml` or `content.xml` to ensure the new image is referenced
    const parser = new xml2js.Parser();
    const builder = new xml2js.Builder();
    const stylesXml = fs.readFileSync(stylesXmlPath, 'utf-8');
    const xmlData = await parser.parseStringPromise(stylesXml);

    // Look for the `<draw:fill-image>` or `<draw:frame>` element referencing the background image
    const drawImageElements = xmlData['office:document-styles']?.['office:styles']?.[0]?.['style:style'];
    if (drawImageElements) {
        for (const style of drawImageElements) {
            if (style['style:graphic-properties']?.[0]?.['$']?.['draw:fill-image-name']) {
                // Update the `draw:fill-image-name` if necessary
                style['style:graphic-properties'][0]['$']['draw:fill-image-name'] = newImageName;
                break;
            }
        }
    }

    const updatedXml = builder.buildObject(xmlData);
    fs.writeFileSync(stylesXmlPath, updatedXml);

    // Re-zip the folder to create the new ODT file
    const newZip = new AdmZip();
    newZip.addLocalFolder(tempDir);
    newZip.writeZip(outputPath);

    // Cleanup temporary files
    //fs.rmSync(tempDir, { recursive: true, force: true });

    console.log(`New ODT saved at ${outputPath}`);
}

app.post('/upload', upload.single('odt'), async (req, res) => {
    const odtPath = req.file.path;
    const newImagePath = req.body.newImagePath; // Path to the replacement image
    const outputPath = path.join('outputs', `output_${Date.now()}.odt`);

    try {
        await replaceBackgroundImage(odtPath, newImagePath, outputPath);
        res.download(outputPath, (err) => {
            if (err) console.error(err);
            fs.unlinkSync(odtPath); // Cleanup uploaded file
        });
    } catch (error) {
        console.error(error);
        res.status(500).send(`Error processing ODT file: ${error.message}`);
    }
});

app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
