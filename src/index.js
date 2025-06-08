import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';

import { ThemeProvider, createTheme } from '@mui/material/styles';
import CssBaseline from '@mui/material/CssBaseline';

/* —— EYP Austria colourway —— */
const theme = createTheme({
  palette: {
    mode: 'light',
    primary:   { main: '#003399' }, // EYP blue :contentReference[oaicite:0]{index=0}
    secondary: { main: '#FFE600' }, // EYP yellow :contentReference[oaicite:1]{index=1}
    background:{ default: '#f5f5f5' }
  }
});

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(
  <ThemeProvider theme={theme}>
    <CssBaseline />
    <App />
  </ThemeProvider>
);
