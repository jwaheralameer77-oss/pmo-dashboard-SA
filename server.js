const express = require('express');
const { spawn } = require('child_process');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Serve static files
app.use(express.static(path.join(__dirname, 'static')));
app.use('/static', express.static(path.join(__dirname, 'static')));

// Health check
app.get('/health', (req, res) => {
    res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Start Python Flask server
let pythonProcess = null;
let flaskReady = false;

function startPythonServer() {
    console.log('Starting Python Flask server...');
    
    const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';
    pythonProcess = spawn(pythonCmd, ['wsgi.py'], {
        cwd: __dirname,
        stdio: 'pipe',
        env: { ...process.env, PORT: '5000' }
    });
    
    pythonProcess.stdout.on('data', (data) => {
        const output = data.toString();
        console.log(`Python: ${output}`);
        if (output.includes('Running on')) {
            flaskReady = true;
            console.log('Flask server is ready!');
        }
    });
    
    pythonProcess.stderr.on('data', (data) => {
        console.error(`Python: ${data}`);
    });
    
    pythonProcess.on('close', (code) => {
        console.log(`Python process exited with code ${code}`);
        flaskReady = false;
        setTimeout(startPythonServer, 3000);
    });
}

// Start Flask
startPythonServer();

// Wait for Flask to be ready
app.use((req, res, next) => {
    if (!flaskReady && req.path !== '/health') {
        return res.status(503).json({ 
            error: 'Server starting',
            message: 'Please wait a few seconds and refresh'
        });
    }
    next();
});

// Proxy to Flask
const httpProxy = require('http-proxy');
const proxy = httpProxy.createProxyServer({
    target: 'http://127.0.0.1:5000',
    changeOrigin: true,
    timeout: 30000,
    proxyTimeout: 30000
});

proxy.on('error', (err, req, res) => {
    console.error('Proxy error:', err.message);
    if (!res.headersSent) {
        res.status(503).json({ 
            error: 'Service temporarily unavailable',
            message: 'Server is starting, please try again'
        });
    }
});

app.all('*', (req, res) => {
    proxy.web(req, res);
});

// Start server
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Node.js server running on port ${PORT}`);
});

// Graceful shutdown
process.on('SIGTERM', () => {
    console.log('Shutting down...');
    if (pythonProcess) pythonProcess.kill();
    process.exit(0);
});
