const express = require('express');
const http = require('http');

const app = express();
const server = http.createServer(app);
const PORT = 3000;

app.get('/', (req, res) => {
    res.send('Hello!');
});

server.listen(PORT, () => {
    console.log(`Test server on http://localhost:${PORT}`);
});

console.log('Script finished executing');
