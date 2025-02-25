const express = require('express');
const fs = require('fs').promises;
const path = require('path');
const mammoth = require('mammoth');

const router = express.Router();

// Helper function to get absolute file path (prevents path traversal attacks)
const getFilePath = (fileName) => path.join(__dirname, '..', 'storage', path.basename(fileName));

// Helper function to check if a file is a Word file
const isWordFile = (fileName) => path.extname(fileName).toLowerCase() === '.docx';

router.get('/read', async (req, res) => {
  const fileName = req.query.fileName;

  if (!fileName) {
    return res.status(400).json({ error: 'File name is required' });
  }

  // Check if the file is a Word file
  if (!isWordFile(fileName)) {
    return res.status(400).json({ error: 'Only Word (.docx) files are allowed' });
  }

  try {
    const filePath = getFilePath(fileName);
    console.log("File Path:", filePath);  // Debug: Print the file path

    // Check if the file exists at the given path
    await fs.access(filePath);  // This checks if the file exists

    // Read the Word file using mammoth
    const data = await fs.readFile(filePath);
    mammoth.extractRawText({ buffer: data })
      .then((result) => {
        // Return the content of the Word file as plain text
        res.json({ content: result.value });
      })
      .catch((err) => {
        res.status(500).json({ error: 'Error extracting text from the Word file' });
      });

  } catch (err) {
    console.error("File not found error:", err);  // Debug: Print error details
    res.status(404).json({ error: 'File not found' });
  }
});

router.post('/append', async (req, res) => {
  const { fileName, content } = req.body;

  // Check if the file is a Word file
  if (!fileName || !isWordFile(fileName)) {
    return res.status(400).json({ error: 'Only Word (.docx) files are allowed' });
  }

  try {
    await fs.appendFile(getFilePath(fileName), content, 'utf8');
    res.json({ message: 'Content appended successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.put('/rename', async (req, res) => {
  const { oldName, newName } = req.body;

  // Check if both old and new names are provided and that the new name is a Word file
  if (!oldName || !newName) {
    return res.status(400).json({ error: 'Both old and new file names are required' });
  }

  if (!isWordFile(newName)) {
    return res.status(400).json({ error: 'New file name must be a Word (.docx) file' });
  }

  const oldFilePath = getFilePath(oldName);
  const newFilePath = getFilePath(newName);

  try {
    await fs.access(oldFilePath);
    await fs.rename(oldFilePath, newFilePath);
    res.json({ message: 'File renamed successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.post('/write', async (req, res) => {
  const { fileName, content } = req.body;

  // Check if the file is a Word file
  if (!fileName || !isWordFile(fileName)) {
    return res.status(400).json({ error: 'Only Word (.docx) files are allowed' });
  }

  try {
    await fs.writeFile(getFilePath(fileName), content, 'utf8');
    res.json({ message: 'File written successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.delete('/delete', async (req, res) => {
  const fileName = req.query.fileName;

  // Check if the file is a Word file
  if (!fileName || !isWordFile(fileName)) {
    return res.status(400).json({ error: 'Only Word (.docx) files are allowed' });
  }

  try {
    await fs.unlink(getFilePath(fileName));
    res.json({ message: 'File deleted successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// router.post('/create-dir', async (req, res) => {
//   const { dirName } = req.body;

//   if (!dirName) {
//     return res.status(400).json({ error: 'Directory name is required' });
//   }

//   const dirPath = getFilePath(dirName);

//   try {
//     await fs.mkdir(dirPath, { recursive: true }); // Creates nested directories if needed
//     res.json({ message: 'Directory created successfully' });
//   } catch (err) {
//     res.status(500).json({ error: err.message });
//   }
// });


// router.delete('/delete-dir', async (req, res) => {
//   const { dirName } = req.query;

//   if (!dirName) {
//     return res.status(400).json({ error: 'Directory name is required' });
//   }

//   const dirPath = getFilePath(dirName);

//   try {
//     await fs.rm(dirPath, { recursive: true, force: true }); // Deletes even if it's not empty
//     res.json({ message: 'Directory deleted successfully' });
//   } catch (err) {
//     res.status(500).json({ error: err.message });
//   }
// });


module.exports = router;
