import { useEffect, useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Container,
  Box,
  Typography,
  CircularProgress,
  Autocomplete,
  TextField,
  Chip,
  Paper,
  Tabs,
  Tab,
  Accordion,
  AccordionSummary,
  AccordionDetails,
  Button,
    Dialog, DialogTitle, DialogContent, DialogActions,
  FormControlLabel, Checkbox
} from "@mui/material";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import { DataGrid, GridToolbar } from "@mui/x-data-grid";
import { groupBy } from "lodash";
import SearchIcon from '@mui/icons-material/Search';   // add at top
import jsPDF from 'jspdf';
import 'jspdf-autotable';
import PictureAsPdfIcon from '@mui/icons-material/PictureAsPdf';
import autoTable from 'jspdf-autotable';
/**
 * ------------------------------------------------------------------
 * 0  | CONFIG
 * ------------------------------------------------------------------
 */
const SPREADSHEET_ID = "1eSsGajvwtzHQFCpYaE_Rhqla7-9KVTktwE5QSHfh4-A";
const XLSX_URL = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=xlsx`;

const LIGHT_BLUE = "#E3F2FD"; // background for rows that contain the tag "blue"
const FALLBACK_BG = "#e0e0e0"; // neutral when no colour is supplied in Config
// Helper – safe trim
const trim = (v) => (typeof v === "string" ? v.trim() : v);

/**
 * ------------------------------------------------------------------
 * 1  | Helpers
 * ------------------------------------------------------------------
 */
const title = (s) =>
  trim(s).replace(/\b\w+\b/g, (word) => {
    const lower = word.toLowerCase();
    if (lower === "it") return "IT";                     // special case
    if (lower === "nc") return "NC";                     // special case
    if (lower === "ic") return "IC";                     // special case
    if (lower === "pr") return "PR";                     // special case
    if (lower === "hr") return "HR";                     // special case
    if (lower === "ue") return "UE";                     // special case



    return lower.charAt(0).toUpperCase() + lower.slice(1);
  });
const norm = (s) => trim(s).toLowerCase();          // unify case/space
const hex   = (c) => c.startsWith("#") ? c : `#${c}`; // add “#” if missing

async function loadWorkbook(url) {
  const buf = await fetch(url).then((r) => r.arrayBuffer());
  return XLSX.read(buf, { type: "array" });
}

function sheetToRows(sheet) {
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const headerIdx = raw.findIndex((r) => r.some((c) => String(c).toLowerCase().includes("measure")));
  const headers = raw[headerIdx].map((h) => trim(h));
  const dataRows = raw.slice(headerIdx + 1).filter((r) => r.some((c) => trim(c) !== ""));
  return dataRows.map((row, i) => {
    const obj = { id: i };
    headers.forEach((h, idx) => (obj[h] = trim(row[idx])));
    return obj;
  });
}
const filteredRows = (rowsArr) =>
  rowsArr.filter((r) => {
    /* same three checks you have in useMemo,
       just replace rows with rowsArr          */
  });

/**
 * ------------------------------------------------------------------
 * 2  | Component
 * ------------------------------------------------------------------
 */
export default function EnhancedDashboard() {
  const [tabs, setTabs] = useState([]);
  const [rowsByTab, setRowsByTab] = useState({});
  const [tabIdx, setTabIdx] = useState(0);
  const [loading, setLoading] = useState(true);

  const [portfolios, setPortfolios] = useState([]);
  const [portfolioOptions, setPortfolioOptions] = useState([]);
  const [portfolioColours, setPortfolioColours] = useState(new Map()); // A‑name ➜ D‑hex

  const [tags, setTags] = useState([]);
  const [tagOptions, setTagOptions] = useState([]);
  const [tagColours, setTagColours] = useState(new Map()); // C‑name ➜ B‑hex
  const [query, setQuery] = useState('');
  
const [openExport, setOpenExport] = useState(false);
const [inclFilters, setInclFilters] = useState(true);   // default = honour filters

  const hasSame = (arr, raw) =>
  arr.some((x) => norm(x) === norm(raw));

  const togglePortfolio = useCallback(
  (rawLabel) =>
    setPortfolios((prev) =>
        hasSame(prev, rawLabel)
        ? prev.filter((p) => norm(p) !== norm(rawLabel))
        : [...prev, rawLabel]
    ),
  []
);

const toggleTag = useCallback(
  (rawLabel) =>
    setTags((prev) =>
        hasSame(prev, rawLabel)
        ? prev.filter((t) => norm(t) !== norm(rawLabel))
        : [...prev, rawLabel]
    ),
  []
);

function handleExport(applyFilters) {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
  const pageWidth = doc.internal.pageSize.getWidth();

  tabs.forEach((tabName, idx) => {
    // 1.  collect rows for this tab
    const allRows    = rowsByTab[tabName];
    const tableRows  = applyFilters ? filteredRows(allRows) : allRows;  // reuse your same filter fn
    if (!tableRows.length) return;                                      // skip empty pillar

    // 2.  work out the columns once
    if (idx === 0) doc.setFontSize(18);
    doc.text(tabName, 40, idx === 0 ? 60 : doc.internal.pageSize.getHeight() - 740);

    const tableCols = columns.map((c) => ({ header: c.headerName, dataKey: c.field }));

    autoTable(doc, {
      startY: idx === 0 ? 80 : undefined,   // after the pillar header
      head: [tableCols.map((c) => c.header)],
      body: tableRows.map((r) => tableCols.map((c) => r[c.dataKey])),
      styles: { fontSize: 8, overflow: 'linebreak', cellWidth: 'wrap' },
      theme: 'grid',
      margin: { left: 40, right: 40 },
      didDrawPage: (d) => {
        // page footer
        doc.setFontSize(10);
        doc.text(`Page ${doc.getNumberOfPages()}`, pageWidth - 60, doc.internal.pageSize.getHeight() - 20);
      },
      addPageContent: idx !== 0,            // create new page per pillar
      showHead: 'everyPage'
    });

    if (idx < tabs.length - 1) doc.addPage();   // make space for next pillar
  });

  doc.save(applyFilters ? 'strategy_filtered.pdf' : 'strategy_all.pdf');
}
  // ---------------------------------------------------------------
  // 2.1 | Load workbook & derive look‑ups
  // ---------------------------------------------------------------
  useEffect(() => {
    loadWorkbook(XLSX_URL)
      .then((wb) => {
        /**
         * Config sheet layout (row 1 is just explanatory text – skip it):
         *  A | Portfolio name
         *  B | Tag colour (hex)
         *  C | Tag name
         *  D | Portfolio colour (hex)
         */
        const cfg = XLSX.utils.sheet_to_json(wb.Sheets.Config, { header: 1, defval: "" });
        const [, ...cfgRows] = cfg; // skip first row (metadata)

        const portSet = new Set();
        const tagSet = new Set();
        const portColourMap = new Map();
        const tagColourMap = new Map();

        cfgRows.forEach((r) => {
          const tagName = trim(r[0]); // col A
          const tagColourHex = trim(r[1]);  // col B
          const portfolioName = trim(r[2]);       // col C
          const portfolioColourHex = trim(r[3]); // col D          // Skip empty rows
          console.log(tagName, portfolioName, tagColourHex, portfolioColourHex);
          if (portfolioName) {
            const key = norm(portfolioName);
            portSet.add(portfolioName);
            if (portfolioColourHex) portColourMap.set(key, hex(portfolioColourHex));
          }
          if (tagName) {
            const key = norm(tagName);
            tagSet.add(tagName);
            if (tagColourHex) tagColourMap.set(key, hex(tagColourHex));
          }
        });

        setPortfolioOptions([...portSet].sort());
        setTagOptions([...tagSet].sort());
        setPortfolioColours(portColourMap);
        setTagColours(tagColourMap);

        // Parse data sheets (skip Introduction + Config)
        const ignore = new Set(["Introduction", "Config"]);
        const sheets = wb.SheetNames.filter((n) => !ignore.has(n));
        const map = Object.fromEntries(sheets.map((name) => [name, sheetToRows(wb.Sheets[name])]));
        setTabs(sheets);
        setRowsByTab(map);
        setLoading(false);
      })
      .catch((err) => {
        console.error(err);
        alert("Failed to load – please check sharing permissions.");
      });
  }, []);

  const rows = tabs.length ? rowsByTab[tabs[tabIdx]] : [];

  // ---------------------------------------------------------------
  // 2.2 | Filtering (portfolio + tags chips)
  // ---------------------------------------------------------------
  const filtered = useMemo(() => {
    return rows.filter((r) => {
    const rowPortfolios = String(r.Portfolio)      // ← singular
        .split(",")
        .map((p) => norm(p)); 
    if (
      portfolios.length &&                         // at least one chip picked
      !portfolios.map(norm).some((p) => rowPortfolios.includes(p))
    ) return false;
      const rowTags = String(r.Tags)
        .split(",")
        .map((t) => trim(t.toLowerCase()));

      if (tags.length && !tags.map((t) => t.toLowerCase()).some((t) => rowTags.includes(t))) return false;
            
      if (
        query &&
          !Object.values(r)                       
            .some((v) =>
              String(v).toLowerCase().includes(query.toLowerCase())
            )
        ) return false;
      return true;
    });
  }, [rows, portfolios, tags, query]);

  // ---------------------------------------------------------------
  // 2.3 | Grouping – extract first two numeric levels
  // ---------------------------------------------------------------
  const grouped = useMemo(() => {
    return groupBy(filtered, (r) => {
      const tryFields = [r.Category, r.Subcategory, r.Measures];
      for (let fld of tryFields) {
        if (!fld) continue;
        const txt = trim(fld);
        const m = txt.match(/^(\d+)\.(\d+)/); // capture 1st two numeric levels
        if (m) return `${m[1]}.${m[2]}`;      }
      return "—"; // bucket for rows with no numeric prefix
    });
  }, [filtered]);

  // ---------------------------------------------------------------
  // 2.4 | Dynamic row height for wrapped text (Measures)
  // ---------------------------------------------------------------
  const getRowHeight = useCallback((params) => {
    const txt = String(params.model.Measures ?? "");
    const charsPerLine = 45;
    const lines = Math.ceil(txt.length / charsPerLine);
    const base = 32; // baseline height
    const lineHeight = 20;
    return base + lines * lineHeight;
  }, []);

  // ---------------------------------------------------------------
  // 2.5 | DataGrid column definitions
  // ---------------------------------------------------------------
  const STATUS_BG = {
  'not started':'#e6e6e6',
  started:'#ffd360',
  maintained:'#bfe1f6',
  delayed:'#ff9689',
  'on track':'#d4edbc',
  'nearly completed':'#98d55e',

  abandoned:'#3d3d3d',
  completed:'#11734b',

};

  const columns = useMemo(() => {
    if (!rows.length) return [];

    return Object.keys(rows[0])
      .filter((f) => !["id", "Category", "Materials"].includes(f) && !/Update/i.test(f))
      .map((field) => {
        const col = {
          field,
          headerName: field,
          minWidth: 140,
          flex: 1,
        };

        // Portfolio ➜ coloured chip
        if (field === 'Portfolio') {
            col.renderCell = (params) => {
              const ports = String(params.value)
                .split(',')
                .map((p) => norm(p))
                .filter(Boolean);

              return (
                <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 0.5 }}>
                  {ports.map((p) => (
                    <Chip
                      key={p}
                      label={title(p)}
                      clickable
                      onClick={() => togglePortfolio(p)}
                      size="small"
                      sx={{
                        bgcolor: portfolioColours.get(p) || FALLBACK_BG,
                        color: '#000',
                      }}
                    />
                  ))}
                </Box>
              );
            };
            col.minWidth = 260;
          }

        // Tags ➜ coloured chips list
        if (field === "Tags") {
          col.renderCell = (params) => {
          const tagsArr = String(params.value)
            .split(",")
            .map((t) => norm(t))
            .filter(Boolean);
            return (
              <Box sx={{ display: "flex", flexWrap: "wrap", gap: 0.5 }}>
                {tagsArr.map((t, i) => (
                <Chip 
                key={t} 
                label={title(t)} 
                size="small"
                clickable
                onClick={() => toggleTag(t)}
                sx={{ bgcolor: tagColours.get(t) || FALLBACK_BG, color: "#000" }} 
                />                
      ))}
              </Box>
            );
          };
          col.minWidth = 260;
        }

        // Measures ➜ wrap & limit width
        if (field === "Measures") {
          col.headerName = "Measures";
          col.renderCell = (params) => (
            <Box sx={{ whiteSpace: "pre-line", lineHeight: 1.4, maxWidth: 420 }}>{String(params.value || "")}</Box>
          );
          col.flex = 2;
          col.minWidth = 300;
        }

        // Yearly status colouring
         if (/^\d{2}\/\d{2}$/.test(field)) {
            col.width = 120;
            col.align = 'center';
            col.headerAlign = 'center';

            // ❶ add a class based on the cell value
            col.cellClassName = (params) => {
              const key = String(params.value || '').toLowerCase();
              // replace spaces so we end up with:  status-on-track / status-delayed …
              return `status-${key.replace(/\s+/g, '-')}`;
            };
          }

        return col;
      });
  }, [rows, portfolioColours, tagColours]);

  /**
   * ------------------------------------------------------------------
   * 3 | Render
   * ------------------------------------------------------------------
   */

  const hasBlueTag = (row) =>
    String(row.Tags)
      .split(",")
      .map((t) => trim(t.toLowerCase()))
      .includes("blue");

  return (
      <Container
        maxWidth={false}   
        sx={{
          mx: '0%',
          pt: 4,            
        }}
      >      
      <Typography variant="h4" gutterBottom>
        National Strategy Dashboard
      </Typography>
       {/* <Button
          variant="outlined"
          size="small"
          startIcon={<PictureAsPdfIcon />}
          onClick={() => setOpenExport(true)}
          sx={{ml: 'auto'}}
        >
          Export PDF
        </Button> */}
        <Dialog open={openExport} onClose={() => setOpenExport(false)} maxWidth="xs" fullWidth>
          <DialogTitle>Export options</DialogTitle>
          <DialogContent>
            <FormControlLabel
              control={
                <Checkbox
                  checked={inclFilters}
                  onChange={(e) => setInclFilters(e.target.checked)}
                />
              }
              label="Apply current filters (Portfolios, Tags, Search)"
            />
          </DialogContent>
          <DialogActions>
            <Button onClick={() => setOpenExport(false)}>Cancel</Button>
            <Button variant="contained" onClick={() => { handleExport(inclFilters); setOpenExport(false); }}>
              Generate PDF
            </Button>
          </DialogActions>
        </Dialog>
      {/* Sheet Tabs */}
      {!loading && (
        <Tabs
          value={tabIdx}
          onChange={(_, v) => setTabIdx(v)}
          sx={{ mb: 3 }}
          variant="scrollable"
          allowScrollButtonsMobile
        >
          {tabs.map((n) => (
            <Tab key={n} label={n} />
          ))}
        </Tabs>
      )}

      {loading ? (
        <Box sx={{ textAlign: "center", mt: 4 }}>
          <CircularProgress />
        </Box>
      ) : (
        <>
          {/* Filters */}
          <Paper
            elevation={0}
            sx={{
              p: 2,
              mb: 3,
              display: "flex",
              gap: 2,
              flexWrap: "wrap",
              alignItems: "center",
              bgcolor: "#fdfdfd",
              border: "1px solid #e0e0e0",
              borderRadius: 2,
            }}
          >
            {/* Portfolio filter */}
            <Autocomplete
              multiple
              size="small"
              options={portfolioOptions}
              value={portfolios}
              onChange={(_, v) =>
                setPortfolios(
                  v.filter(
                    (val, i, arr) =>
                      arr.findIndex((x) => norm(x) === norm(val)) === i  // keep first hit
                  )
                )
              }              
              sx={{ minWidth: 240 }}
              renderTags={() => null}
              renderInput={(params) => (
                <TextField {...params} label="Portfolios" placeholder="Select…" />
              )}
            />
            


            {/* Tag filter */}
            <Autocomplete
              multiple
              size="small"
              options={tagOptions}
              value={tags}
              onChange={(_, v) =>
                  setTags(
                    v.filter(
                      (val, i, arr) =>
                        arr.findIndex((x) => norm(x) === norm(val)) === i  // keep first hit
                    )
                  )
                }              
                sx={{ minWidth: 240 }}
                renderTags={() => null}                
                renderInput={(params) => (
              <TextField {...params} label="Tags" placeholder="Select…" />
  )}
            />
            <Box
              sx={{
                mt: 1.5,
                display: 'flex',
                flexWrap: 'wrap',
                gap: 0.5,
              }}
            >
              {/* Portfolios first (greenish?) */}
              {portfolios.map((p) => (
                <Chip
                  key={`port-${p}`}
                  label={title(p)}
                  onDelete={() =>
                    setPortfolios((prev) => prev.filter((name) => name !== p))
                  }
                  sx={{
                    bgcolor: portfolioColours.get(norm(p)) || FALLBACK_BG,
                    color: '#000',
                  }}
                />
              ))}

              {/* Tags next (blue-ish?) */}
              {tags.map((t) => (
                <Chip
                  key={`tag-${t}`}
                  label={title(t)}
                  onDelete={() =>
                    setTags((prev) => prev.filter((name) => name !== t))
                  }
                  sx={{
                    bgcolor: tagColours.get(norm(t)) || FALLBACK_BG,
                    color: '#000',
                  }}
                />
              ))}
            </Box>
            <TextField
              size="small"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search…"
              sx={{ ml: 'auto', width: 260 }}
              InputProps={{
                startAdornment: (
                  <SearchIcon sx={{ mr: 1, color: 'text.secondary' }} />
                ),
              }}
            />
          </Paper>

          {/* Tables by Category */}
          {Object.entries(grouped).map(([cat, items]) => {
            const fullCategory = items[0].Category || `Category ${cat}`;
            return (
              <Accordion key={cat} defaultExpanded sx={{ mb: 3 }}>
                <AccordionSummary expandIcon={<ExpandMoreIcon />}>
                  <Typography variant="h6">{fullCategory}</Typography>
                </AccordionSummary>
                <AccordionDetails>
                  <Box sx={{ width: "100%" }}>
                    <DataGrid
                      getRowHeight={getRowHeight}
                      rows={items}
                      columns={columns}
                      disableRowSelectionOnClick
                      components={{ Toolbar: GridToolbar }}
                      density="compact"
                      getRowClassName={(params) => (hasBlueTag(params.row) ? "row-blue" : "")}
                      sx={{
                        "& .row-blue": { bgcolor: LIGHT_BLUE },
                        "& .MuiDataGrid-cell": { lineHeight: 1.4, whiteSpace: "normal", py: 1 },
                        '& .status-not-started'  : { bgcolor: STATUS_BG['not started'] },
                        '& .status-started'      : { bgcolor: STATUS_BG.started },
                        '& .status-maintained'   : { bgcolor: STATUS_BG.maintained },
                        '& .status-delayed'      : { bgcolor: STATUS_BG.delayed },
                        '& .status-nearly-completed'     : { bgcolor: STATUS_BG['nearly completed'] },
                        '& .status-abandoned'    : { bgcolor: STATUS_BG.abandoned},
                        '& .status-completed'      : { bgcolor: STATUS_BG.completed },
                        '& .status-on-track'      : { bgcolor: STATUS_BG['on track'] },
                      }}
                    />
                  </Box>
                </AccordionDetails>
              </Accordion>
            );
          })}

        </>
      )}
    </Container>
  );
}

