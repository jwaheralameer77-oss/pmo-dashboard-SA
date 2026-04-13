const express = require('express');
const { spawn } = require('child_process');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 10000;

// Serve static files
app.use(express.static(path.join(__dirname, 'static')));
app.use('/static', express.static(path.join(__dirname, 'static')));

// Health check
app.get('/health', (req, res) => {
    res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Start Python Flask server
let flaskProcess = null;
let flaskReady = false;

function startFlask() {
    console.log('Starting Flask server...');
    
    const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';
    
    flaskProcess = spawn(pythonCmd, ['-c', `
import sys
sys.path.insert(0, '.')
from app import app, db
from seed import seed_database

with app.app_context():
    db.create_all()
    seed_database()

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=False)
`], {
        cwd: __dirname,
        stdio: 'pipe'
    });
    
    flaskProcess.stdout.on('data', (data) => {
        const output = data.toString();
        console.log(`Flask: ${output}`);
        if (output.includes('Running on')) {
            flaskReady = true;
            console.log('Flask is ready!');
        }
    });
    
    flaskProcess.stderr.on('data', (data) => {
        console.error(`Flask Error: ${data}`);
    });
    
    flaskProcess.on('close', (code) => {
        console.log(`Flask exited with code ${code}`);
        flaskReady = false;
        setTimeout(startFlask, 3000);
    });
}

// Start Flask
startFlask();

// Wait for Flask
app.use((req, res, next) => {
    if (req.path === '/health') return next();
    if (!flaskReady) {
        return res.status(503).send(`
            <html><body style="text-align:center; padding:50px; font-family:Arial;">
                <h2>⏳ جاري تشغيل الخادم...</h2>
                <p>الرجاء الانتظار 30 ثانية ثم تحديث الصفحة</p>
                <script>setTimeout(() => location.reload(), 5000);</script>
            </body></html>
        `);
    }
    next();
});

// Proxy to Flask
const httpProxy = require('http-proxy');
const proxy = httpProxy.createProxyServer({
    target: 'http://127.0.0.1:5000',
    changeOrigin: true
});

proxy.on('error', (err, req, res) => {
    console.error('Proxy error:', err.message);
    if (!res.headersSent) {
        res.status(503).send('Service temporarily unavailable');
    }
});

app.all('*', (req, res) => {
    proxy.web(req, res);
});

// Start
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Node.js server on port ${PORT}`);
});

process.on('SIGTERM', () => {
    if (flaskProcess) flaskProcess.kill();
    process.exit(0);
});
