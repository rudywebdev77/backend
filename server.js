import app from './app.js';

const PORT = process.env.PORT || 5000;

// Start Server on 0.0.0.0 for LAN & Mobile accessibility
app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server running on port ${PORT} (Bound to 0.0.0.0 for LAN/mobile access)`);
});
