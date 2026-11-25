# Virani & Sorathia wedding website

Shazeena Virani weds Shahzad Sorathia · 29th – 30th January 2027 · Denver, Colorado

A single-page wedding website: opening doors with a tassel and invitation card, our story, the three celebrations (Pithi / Sangeet, Nikkah, Reception), travel and stay, and the RSVP link.

Live at https://shahzadsorathia.github.io/VS-Wedding/

## Folder structure

```
VS-Wedding/
├── index.html              the whole page (styles and scripts included)
├── README.md
├── .nojekyll               tells GitHub Pages to serve files as they are
├── assets/
│   ├── images/
│   │   ├── hero/           opening card layers, back to front:
│   │   │                   sky, mountains-far, mountains-mid, arch,
│   │   │                   wisteria-a, wisteria-b, couple, foreground
│   │   └── story/          bench, proposal, winter
│   ├── video/              save-the-date.mp4 and its poster.jpg
│   ├── audio/              song.mp3 (plays after the tassel is pulled)
│   ├── icons/              browser tab and home-screen icons
│   └── share/              og-image.jpg, the link-preview picture
└── apps-script/            copy of the RSVP form and couple's dashboard
    ├── Code.gs             (runs on Google Apps Script, not here)
    ├── Index.html
    ├── Dashboard.html
    └── README.md
```

`index.html` must stay in the top folder: GitHub Pages serves it as the home page.

Hosted with GitHub Pages. The RSVP form and couple's dashboard run on Google Apps Script; the `apps-script/` folder is a copy for safekeeping. Settings such as the dashboard password live in Apps Script's Script Properties, never in this repository.
