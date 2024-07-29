document.getElementById('uploadForm').addEventListener('submit', function(event) {
    event.preventDefault();

    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    if (!file) {
        document.getElementById('message').innerText = 'Por favor, selecciona un archivo Excel.';
        return;
    }

    const formData = new FormData();
    formData.append('file', file);

    fetch('https://python-api-whatsapp-backend.onrender.com/upload', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            document.getElementById('message').innerText = 'Mensajes enviados correctamente.';
        } else {
            document.getElementById('message').innerText = `Error: ${data.error}`;
        }
    })
    .catch(error => {
        document.getElementById('message').innerText = `Error de red: ${error.message}`;
    });
});
