const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, ImageRun } = require("docx");

// Path to the folder containing images
const folderPath = "C:/Users/harsh/Downloads/Results/Results/ScreenShots/ESIC_Hospital_4th F"; // Replace with your actual folder path

// Function to load images and create the Word document
async function createWordWithImages() {
    // Create a new Word document with metadata (adding creator, title, and description)
    const doc = new Document({
        creator: "Harsh Langeh", // Set the creator metadata
        title: "ESIC Hospital Screenshots",
        description: "A Word document containing images from the ESIC Hospital folder.",
        sections: [] // Initialize sections as an empty array
    });

    // Read all files from the folder
    const files = fs.readdirSync(folderPath);

    for (const file of files) {
        const filePath = path.join(folderPath, file);

        // Ensure the file is an image (basic check for extensions)
        if (/\.(png|jpg|jpeg|gif)$/i.test(file)) {
            // Add the image name as the title
            doc.addSection({
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: file, // Image file name
                                bold: true,
                                size: 32, // Title size
                            }),
                        ],
                        spacing: { after: 300 }, // Add spacing after the title
                    }),
                    // Add the image to the document
                    new Paragraph({
                        children: [
                            new ImageRun({
                                data: fs.readFileSync(filePath),
                                transformation: {
                                    width: 600, // Adjust width
                                    height: 400, // Adjust height
                                },
                            }),
                        ],
                    }),
                ],
            });
        }
    }

    // Save the document as a .docx file
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync("images_with_titles.docx", buffer);

    console.log("Word document created successfully: images_with_titles.docx");
}

// Run the function
createWordWithImages().catch((error) => {
    console.error("Error creating Word document:", error);
});
