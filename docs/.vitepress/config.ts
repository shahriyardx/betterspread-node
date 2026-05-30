import { defineConfig } from "vitepress"

export default defineConfig({
  title: "betterspread",
  description:
    "TypeScript Google Sheets wrapper with typed rows, cell formatting, and Zod schema validation",
  themeConfig: {
    nav: [
      { text: "Guide", link: "/guide/" },
      { text: "API", link: "/api/" },
      {
        text: "GitHub",
        link: "https://github.com/shahriyardx/betterspread-node",
      },
    ],
    sidebar: {
      "/guide/": [
        {
          text: "Guide",
          items: [
            { text: "Getting Started", link: "/guide/" },
            { text: "Rows & Cells", link: "/guide/rows-cells" },
            { text: "Cell Formatting", link: "/guide/formatting" },
            { text: "Zod Validation", link: "/guide/validation" },
            { text: "Formulas", link: "/guide/formulas" },
          ],
        },
      ],
      "/api/": [
        {
          text: "API Reference",
          items: [
            { text: "Connection", link: "/api/connection" },
            { text: "Sheet", link: "/api/sheet" },
            { text: "Tab", link: "/api/tab" },
            { text: "Row", link: "/api/row" },
            { text: "Cell", link: "/api/cell" },
            { text: "Format", link: "/api/format" },
            { text: "Types", link: "/api/types" },
          ],
        },
      ],
    },
    socialLinks: [
      {
        icon: "github",
        link: "https://github.com/shahriyardx/betterspread-node",
      },
    ],
  },
})
