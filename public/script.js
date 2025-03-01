const apiUrl = 'http://localhost:3000/api';

// Helper function to check if the file name ends with '.docx'
function isWordFile(fileName) {
  return fileName.toLowerCase().endsWith('.docx');
}

// Read file
async function readFile() {
  const fileName = document.getElementById('readFileName').value;
  if (!fileName) {
    alert('File name is required');
    return;
  }
  if (!isWordFile(fileName)) {
    alert('Only Word (.docx) files are allowed');
    return;
  }
  
 
  const response = await fetch(`${apiUrl}/read?fileName=${fileName}`);
  const data = await response.json();
  document.getElementById('readContent').textContent = data.content || data.error;
}

// Write file
async function writeFile() {
  const fileName = document.getElementById('writeFileName').value;
  const content = document.getElementById('writeContent').value;
  if (!fileName) {
    alert('File name is required');
    return;
  }

  if (!isWordFile(fileName)) {
    alert('Only Word (.docx) files are allowed');
    return;
  }

  await fetch(`${apiUrl}/write`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fileName, content })
  });
}

// Append file
async function appendFile() {
  const fileName = document.getElementById('appendFileName').value;
  const content = document.getElementById('appendContent').value;
  if (!fileName) {
    alert('File name is required');
    return;
  }

  if (!isWordFile(fileName)) {
    alert('Only Word (.docx) files are allowed');
    return;
  }

  await fetch(`${apiUrl}/append`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fileName, content })
  });
}

// Delete file
async function deleteFile() {
  const fileName = document.getElementById('deleteFileName').value;
  if (!fileName) {
    alert('File name is required');
    return;
  }

  if (!isWordFile(fileName)) {
    alert('Only Word (.docx) files are allowed');
    return;
  }

  await fetch(`${apiUrl}/delete?fileName=${fileName}`, { method: 'DELETE' });
}

// Rename file
async function renameFile() {
  const oldName = document.getElementById('oldFileName').value;
  const newName = document.getElementById('newFileName').value;
  const fileName = document.getElementById('appendFileName').value;

  if (!fileName) {
    alert('File name is required');
    return;
  }


  if (!isWordFile(oldName) || !isWordFile(newName)) {
    alert('Only Word (.docx) files are allowed');
    return;
  }

  await fetch(`${apiUrl}/rename`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ oldName, newName })
  });
}

// Create directory
async function createDirectory() {
  const dirName = document.getElementById('createDirName').value;
  await fetch(`${apiUrl}/create-dir`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ dirName })
  });
}

// Delete directory
async function deleteDirectory() {
  const dirName = document.getElementById('deleteDirName').value;
  await fetch(`${apiUrl}/delete-dir?dirName=${dirName}`, { method: 'DELETE' });
}
