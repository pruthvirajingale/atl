const express = require('express');
const path = require('path');
const app = express();

// Make sure this path exists!
app.use(express.static(path.join(__dirname, 'attendance-server')));

app.listen(3000, () => console.log('Server running on port 3000'));