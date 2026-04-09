# PowerDocu Generated Output Structure

This document describes the folder and file structure that PowerDocu generates when documenting a Power Platform solution. It is intended as a reference for AI-assisted review and analysis of the generated output.

## Root Folder Layout

When PowerDocu processes a solution, it creates the following structure:

```
Solution <SolutionName>/
│
├── FlowDoc <FlowName>/              -- one folder per Cloud Flow
├── AppDoc <AppName>/                 -- one folder per Canvas App
├── AgentDoc <AgentName>/             -- one folder per Agent / Copilot
├── MDADoc <AppName>/                 -- one folder per Model-Driven App
├── AIModelDoc <ModelName>/           -- one folder per AI Model
├── BPFDoc <BPFName>/                 -- one folder per Business Process Flow
├── DesktopFlowDoc <FlowName>/       -- one folder per Desktop Flow
│
├── Dataverse/                        -- Form wireframe SVGs per Dataverse table
├── Resources/                        -- Security role access level icons
├── WebResources/                     -- Extracted web resource files + detail pages
│
├── solution-<name>.html / .md        -- Solution overview
├── Solution - <name>.docx            -- Solution overview (Word)
├── webresources-<name>.html / .md    -- Web Resources documentation
├── WebResources - <name>.docx        -- Web Resources (Word)
├── dataverse.png / dataverse.svg     -- Dataverse entity relationship diagram
├── solution-components.png / .svg    -- Cross-component relationship graph
├── style.css                         -- Shared stylesheet for HTML output
└── .order                            -- Ordering file for wiki-style renderers
```

Each component in the solution gets its own subfolder. The number of folders depends on how many components exist in the solution.

## Output Formats

Each component is documented in up to three formats (configurable):

| Format | Extension | Notes |
|--------|-----------|-------|
| **Word** | `.docx` | Single file per component with all sections |
| **HTML** | `.html` | Multiple interlinked pages with navigation sidebar, plus `style.css` |
| **Markdown** | `.md` | Multiple interlinked files, compatible with Azure DevOps wiki |

All three formats can be generated simultaneously. Diagrams (PNG + SVG) are always generated regardless of format selection.

## Component Folder Details

### FlowDoc \<Name\>/ — Cloud Flow

Documents a Power Automate cloud flow: metadata, trigger configuration, action steps, connection references, and variables.

| File | Purpose |
|------|---------|
| `flow.png` / `flow.svg` | Top-level flow diagram |
| `flow-detailed.png` / `flow-detailed.svg` | Detailed flow diagram with all subactions expanded |
| `<name>.docx` | Word document |
| `index-<name>.html` / `.md` | Overview page |
| `connections-<name>.html` / `.md` | Connection references |
| `variables-<name>.html` / `.md` | Variables with cross-references to actions |
| `triggersactions-<name>.html` / `.md` | Trigger and action details |
| `actions/` (HTML only) | One HTML page per action (`Trigger.html`, `<ActionName>.html`) |

### AppDoc \<Name\>/ — Canvas App

Documents a Power Apps canvas app: screens, controls, properties, variables, data sources, and resources.

| File | Purpose |
|------|---------|
| `ScreenNavigation.png` / `.svg` | Screen-to-screen navigation diagram |
| `<name>.docx` | Word document (plus per-screen Word docs) |
| `index-<name>.html` / `.md` | Overview page |
| `appdetails-<name>.html` / `.md` | App properties and preview flags |
| `variables-<name>.html` / `.md` | Global variables, context variables, collections |
| `datasources-<name>.html` / `.md` | Data sources and table definitions |
| `resources-<name>.html` / `.md` | App resources (images, media) |
| `controls-<name>.html` / `.md` | Controls overview with hierarchical tree |
| `screen-<screenName>-<name>.html` / `.md` | Per-screen detailed control properties |
| `datasource-<dsName>-<name>.html` / `.md` | Per-datasource detail pages |
| `resources/` | Control type icons and app logo |

### AgentDoc \<Name\>/ — Agent / Copilot

Documents a Copilot Studio agent: knowledge sources, tools (flows), connected agents, topics with dialog flows, entities, and variables.

| File | Purpose |
|------|---------|
| `topic-dataflow.png` / `.svg` | Inter-topic data flow diagram |
| `<name>.docx` | Word document |
| `index-<name>.html` / `.md` | Overview page |
| `knowledge-<name>.html` / `.md` | Knowledge sources |
| `tools-<name>.html` / `.md` | Tools (flows, connectors) |
| `agents-<name>.html` / `.md` | Connected agents |
| `topics-<name>.html` / `.md` | Topics overview |
| `entities-<name>.html` / `.md` | Custom entities |
| `variables-<name>.html` / `.md` | Variables with global usage cross-references |
| `channels-<name>.html` / `.md` | Channel configuration |
| `settings-<name>.html` / `.md` | AI and security settings |
| `Topics/` | Per-topic flow diagrams (`<topic>-detailed.png/svg`) and detail pages |
| `Knowledge/` | Per-knowledge-source detail pages |

### MDADoc \<Name\>/ — Model-Driven App

Documents a model-driven app: site map navigation, security roles, referenced tables, views, and app settings.

| File | Purpose |
|------|---------|
| `sitemap-<name>.svg` | Site map navigation visualization |
| `<name>.docx` | Word document |
| `mda-<name>.html` / `.md` | Documentation page |

### AIModelDoc \<Name\>/ — AI Model

Documents a custom AI model: name, type, instructions/prompt, inputs, and model response configuration.

| File | Purpose |
|------|---------|
| `<name>.docx` | Word document |
| `aimodel-<name>.html` / `.md` | Documentation page |

### BPFDoc \<Name\>/ — Business Process Flow

Documents a business process flow: primary entity, stages with steps and required fields, cross-entity stages, and properties.

| File | Purpose |
|------|---------|
| `<name>.docx` | Word document |
| `bpf-<name>.html` / `.md` | Documentation page |

### DesktopFlowDoc \<Name\>/ — Desktop Flow

Documents a Power Automate desktop flow: Robin script actions, subflows, variables, control flow blocks, modules, and connectors.

| File | Purpose |
|------|---------|
| `desktopflow-detailed.png` / `.svg` | Main flow diagram |
| `desktopflow-subflow-<name>-detailed.png` / `.svg` | Per-subflow diagrams |
| `<name>.docx` | Word document |
| `desktopflow-<name>.html` / `.md` | Documentation page |

## Solution-Level Documentation

The files at the root of the output folder provide a solution-wide overview.

**Solution Overview** (`solution-<name>.html/.md/.docx`) covers:
- Solution properties: managed/unmanaged status, version, publisher details
- Component statistics: counts per type
- Environment variables, app actions, setting definitions, formula definitions
- All solution components grouped by type: Tables (with columns, forms, views), Security Roles (with privilege matrices), Option Sets, Connectors, Web Resources, Workflows, and more

**Component Relationship Graph** (`solution-components.png/svg`):
Directed graph showing how components reference each other. Node shapes and colors distinguish component types (e.g. hexagons for agents, rectangles for flows/apps, cylinders for tables).

**Dataverse Entity Relationship Diagram** (`dataverse.png/svg`):
Undirected graph showing table-to-table relationships via lookup columns (one-to-many) and many-to-many associations.

**Dataverse/ subfolder**:
SVG wireframe mockups of Dataverse entity forms, showing tabs, sections, and control layout.

**WebResources/ subfolder**:
Extracted web resource files (images, scripts, stylesheets) plus individual detail pages.

## Component Types and Relationships

PowerDocu discovers cross-component relationships and visualizes them in the component relationship graph. The following relationship types exist:

| Source | Target | Relationship | Meaning |
|--------|--------|-------------|---------|
| Agent | Flow | uses as tool | Agent invokes a cloud flow as an action tool |
| Agent | Connector | uses connector | Agent uses a connector via tool connection |
| Agent | Agent | invokes agent | Agent orchestrates another agent |
| Agent | Table | knowledge source | Agent uses a Dataverse table for knowledge |
| Agent | AI Model | uses AI model | Agent uses a custom AI model |
| Flow | Connector | uses connector | Flow has a connection reference to a connector |
| Canvas App | Table | data source | App reads/writes a Dataverse table |
| Canvas App | Data Source | data source | App uses a non-Dataverse external data source |
| Model-Driven App | Table | uses table | App references a table via app components |
| Model-Driven App | Table | navigates to | App references a table via site map navigation |
| Business Process Flow | Table | primary entity | BPF is bound to a primary table |
| Business Process Flow | Table | stage entity | BPF has cross-entity stages on additional tables |
| Desktop Flow | Connector | uses connector | Desktop flow uses a connector |
| Desktop Flow | Environment Variable | uses env variable | Desktop flow references an environment variable |
| Table | Table | lookup | One-to-many relationship via lookup column |
| Table | Table | many-to-many | Many-to-many inter-table relationship |

## Navigation and Cross-Links

HTML and Markdown output includes navigation links between component documentation pages and the solution overview. Each component page links back to the solution overview, and the solution overview links to each component's documentation folder. Data source references in apps link to the corresponding table documentation in the solution overview.
