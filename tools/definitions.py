TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "raw_web_browser",
            "description": "Get, parse and display detailed structure and content of a webpage",
            "parameters": {
                "type": "object",
                "properties": {
                    "url": {"type": "string", "description": "URL of the webpage to analyze"},
                    "show_details": {"type": "boolean", "description": "Whether to include the original HTML code. Default is False", "default": False}
                },
                "required": ["url"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "perform_searchapi_search",
            "description": """Perform search using searchapi.io with support for 40+ search engines and platforms.

SUPPORTED ENGINES:
  Google Family: google, google_images, google_videos, google_news, google_maps, google_shopping, google_flights, google_hotels, google_scholar, google_jobs, google_events, google_trends, google_finance, google_patents, google_lens, google_autocomplete
  
  Other Search Engines: bing, bing_images, bing_videos, yahoo, baidu, yandex, duckduckgo, naver
  
  E-commerce: amazon, shein, walmart, ebay
  
  Social & Ads: youtube, meta_ad_library, linkedin_ad_library, tiktok_ads, tiktok_profile, facebook, instagram
  
  Travel: airbnb, tripadvisor
  
  Apps: google_play, apple_app_store

USAGE EXAMPLES:
  - General web search: engine="google", query="AI news"
  - Image search: engine="google_images", query="sunset"
  - Product search: engine="amazon", query="laptop"
  - Video search: engine="youtube", query="python tutorial"
  - Job search: engine="google_jobs", query="software engineer"
  - News search: engine="google_news", query="technology"
  - Academic search: engine="google_scholar", query="machine learning"
  - Travel search: engine="airbnb", query="Tokyo apartments"
  - Social media: engine="instagram", query="username"
  - Ad research: engine="meta_ad_library", query="brand name"

COMMON PARAMETERS (vary by engine):
  - location: Geographic location for localized results
  - num: Number of results to return
  - page: Page number for pagination
  - language: Language code (e.g., "en", "zh")
  - date_range: Filter by date (e.g., "past_day", "past_week")
""",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "Keywords to search for"},
                    "engine": {
                        "type": "string", 
                        "description": "Search engine to use (default: google)",
                        "enum": [
                            "google", "google_images", "google_videos", "google_news", "google_maps",
                            "google_shopping", "google_flights", "google_hotels", "google_scholar",
                            "google_jobs", "google_events", "google_trends", "google_finance",
                            "google_patents", "google_lens", "google_autocomplete",
                            "bing", "bing_images", "bing_videos", "yahoo", "baidu", "yandex",
                            "duckduckgo", "naver", "amazon", "shein", "walmart", "ebay",
                            "youtube", "meta_ad_library", "linkedin_ad_library", "tiktok_ads",
                            "tiktok_profile", "facebook", "instagram", "airbnb", "tripadvisor",
                            "google_play", "apple_app_store"
                        ],
                        "default": "google"
                    }
                },
                "required": ["query"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "call_weather_api",
            "description": "Get various types of weather information",
            "parameters": {
                "type": "object",
                "properties": {
                    "api_type": {
                        "type": "string",
                        "description": "API type",
                        "enum": ["current", "forecast", "history", "alerts", "marine", "future", "astronomy", "timezone", "sports", "ip", "search"]
                    },
                    "location": {"type": "string", "description": "Location query", "default": "auto:ip"}
                },
                "required": ["api_type"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "search_scira",
            "description": """Perform AI-powered search using Scira API with 4 specialized agents.

SCIRA AI AGENTS:
  1. search: General web search with AI summary and sources
     - Returns AI-analyzed summary with inline citations
     - Includes source URLs and structured results
     - Best for: Research, fact-checking, general questions
  
  2. people: Find and enrich people profiles
     - Aggregates profiles from LinkedIn, GitHub, X (Twitter), personal websites
     - Returns comprehensive bio, social links, and professional info
     - Best for: Finding contact info, researching individuals, networking
  
  3. xsearch: Real-time X (Twitter) search
     - Search all of X or target specific user
     - Returns AI summary with tweet details (likes, replies, media)
     - Best for: Social media monitoring, brand research, trending topics
  
  4. reddit: Search Reddit discussions
     - AI-summarized answers from Reddit threads
     - Includes relevant thread excerpts and URLs
     - Best for: Community opinions, troubleshooting, product reviews

USAGE EXAMPLES:
  - Web search: agent="search", query="Latest AI developments"
  - Find person: agent="people", query="Andrej Karpathy AI researcher"
  - X search: agent="xsearch", query="AI Gateway", username="vercel" (optional)
  - Reddit search: agent="reddit", query="Best practices for Next.js deployment"

RETURN FORMAT:
  All agents return:
  {
    "text": "AI-generated summary with inline citations",
    "sources": ["https://...", "https://..."],
    "results": [...] // Raw results (varies by agent)
  }

VS SEARCHAPI:
  - SearchAPI: Returns raw search results from 40+ engines
  - Scira: Returns AI-analyzed summaries with citations
  - Use SearchAPI when you need raw data
  - Use Scira when you need AI analysis and synthesis
""",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "Search query or question"},
                    "agent": {
                        "type": "string",
                        "description": "Agent type (default: search)",
                        "enum": ["search", "people", "xsearch", "reddit"],
                        "default": "search"
                    },
                    "username": {
                        "type": "string",
                        "description": "(For xsearch agent only) Target specific X/Twitter user. Omit to search all of X."
                    },
                    "timeout": {
                        "type": "integer",
                        "description": "Request timeout in seconds (default: 30)",
                        "default": 30
                    }
                },
                "required": ["query"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "run_terminal_command",
            "description": "Run a terminal command on the local machine and see real-time output",
            "parameters": {
                "type": "object",
                "properties": {
                    "command": {"type": "string", "description": "The shell command to execute"}
                },
                "required": ["command"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "generate_word_document",
            "description": """Create a professional, beautifully styled Word document (.docx) with rich content.

DOCUMENT SETTINGS (document_settings):
  - default_font: font name (e.g. "Calibri", "SimSun")
  - default_font_size: size in pt (e.g. 11)
  - margins: {top, bottom, left, right} in cm
  - header_text / footer_text: text shown on every page
  - page_numbers: true to add page numbers
  - line_spacing: e.g. 1.5
  - orientation: "portrait" or "landscape"

CONTENT BLOCK TYPES (content_structure array):
  1. cover_page: {type:"cover_page", title, subtitle, author, date, logo_path, title_color:"#003366", title_font_size:36, subtitle_font_size:18, subtitle_color:"#666666"}
  2. toc: {type:"toc"} — inserts Table of Contents
  3. heading: {type:"heading", text, level(1-9), format:{alignment, color:"#003366", font_name, font_size}}
  4. paragraph: {type:"paragraph", text(string or rich parts list), format:{alignment:"justify", line_spacing:1.5, first_line_indent:0.75, font_size, color, bold, italic, font_name}}
  5. rich_paragraph: {type:"rich_paragraph", parts:[{text, bold, italic, color, font_size, font_name, underline, superscript, subscript}], format:{alignment}}
  6. bullet_list: {type:"bullet_list", items:["item1", {text:"item2", level:1}]}
  7. numbered_list: {type:"numbered_list", items:["item1", {text:"item2", level:1}]}
  8. table: {type:"table", rows:[["Header1","Header2"],["val1","val2"]], header_bg_color:"#4472C4", header_font_color:"#FFFFFF", stripe_colors:["#FFFFFF","#F2F2F2"], col_widths:[2,3], font_size:10, alignment:"center", style:"Table Grid"}
  9. image: {type:"image", path:"relative/path.png", width:5, alignment:"center", caption:"Figure 1"}
  10. chart: {type:"chart", chart_type:"bar|line|pie|scatter|area|bar_horizontal|grouped_bar|stacked_bar", title, categories:["A","B"], series:[{name:"S1", values:[10,20]}], colors:["#4472C4"], width:8, height:5, show_values:true, show_legend:true, doc_width:6, font_size:11}
  11. quote: {type:"quote", text:"quoted text", author:"Author Name"}
  12. code_block: {type:"code_block", code:"print('hello')", language:"python"}
  13. horizontal_rule: {type:"horizontal_rule"}
  14. page_break: {type:"page_break"}
  15. section_break: {type:"section_break", orientation:"landscape"}
  16. watermark: {type:"watermark", text:"DRAFT"}

CHART TYPES for type "chart":
  - bar / grouped_bar: vertical bar chart (grouped if multiple series)
  - stacked_bar: stacked vertical bars
  - bar_horizontal: horizontal bar chart
  - line: line chart with markers
  - pie: pie chart (uses first series only)
  - scatter: scatter plot
  - area: filled area chart

TIPS FOR BEAUTIFUL DOCUMENTS:
  - Always start with cover_page + toc for formal documents
  - Use consistent heading colors (e.g. "#003366" for all headings)
  - Use charts to visualize data instead of just tables
  - Add horizontal_rule between major sections
  - Use quote blocks for key insights or citations
  - Use rich_paragraph for mixed formatting within a paragraph
  - Set document_settings with page_numbers:true and appropriate margins""",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Path to save the .docx file (relative to workspace)"},
                    "content_structure": {
                        "type": "array",
                        "description": "Array of content blocks. Each block has a 'type' field and type-specific fields as described above.",
                        "items": {"type": "object"}
                    },
                    "document_settings": {
                        "type": "object",
                        "description": "Document-level settings: default_font, default_font_size, margins{top,bottom,left,right}, header_text, footer_text, page_numbers, line_spacing, orientation"
                    }
                },
                "required": ["file_path", "content_structure"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "generate_excel_document",
            "description": """Create a professional, beautifully styled Excel document (.xlsx) with charts, conditional formatting, and advanced features.

WORKBOOK SETTINGS (workbook_settings): optional global settings.

SHEET DEFINITION (each item in sheets_data):
  - name: sheet tab name
  - data: 2D array [[row1], [row2], ...] — first row is treated as header
  - formulas: [{cell:"C11", formula:"=SUM(C2:C10)"}]
  - column_widths: {"A":15, "B":20} or [15, 20, ...] — auto-calculated if omitted
  - row_heights: {"1": 30}
  - merge_cells: ["A1:C1", "D5:F5"]
  - freeze_panes: "A2" — freeze header row
  - auto_filter: "A1:F1" — add filter dropdowns

STYLING:
  - header_style: {bold:true, font_size:11, font_color:"FFFFFF", bg_color:"4472C4", alignment:"center", border:{color:"4472C4", style:"thin"}}
  - data_style: {font_size:10, alignment:"center", wrap_text:true}
  - stripe_colors: ["F2F7FC", "FFFFFF"] — alternating row colors
  - cell_styles: [{range:"A1:C1", bold:true, bg_color:"FFD700"}, {cell:"D5", font_color:"FF0000"}]

CONDITIONAL FORMATTING (conditional_formatting array):
  - color_scale: {range, type:"color_scale", min_color:"FF0000", max_color:"00FF00"}
  - data_bar: {range, type:"data_bar", color:"4472C4"}
  - cell_is: {range, type:"cell_is", operator:"greaterThan|lessThan|between|equal", value:100, font_color, bg_color}
  - icon_set: {range, type:"icon_set", icon_style:"3Arrows|3TrafficLights|4Arrows"}

CHARTS (charts array):
  - type: "bar", "line", "pie", "area", "scatter"
  - title: chart title
  - data_range: {min_col, max_col, min_row, max_row}
  - categories_range: {min_col, min_row, max_row}
  - position: cell position like "E2"
  - width, height: chart dimensions
  - style: chart style number (1-48)
  - x_axis_title, y_axis_title

DATA VALIDATIONS (data_validations array):
  - list: {range, type:"list", formula:"Option1,Option2,Option3"}
  - whole/decimal: {range, type:"whole", min:0, max:100}

PRINT SETTINGS (print_settings):
  - orientation: "landscape" or "portrait"
  - fit_to_page: true

TIPS FOR BEAUTIFUL SPREADSHEETS:
  - Always use header_style with a professional color scheme
  - Use stripe_colors for alternating row colors (easier to read)
  - Add freeze_panes:"A2" to keep headers visible
  - Use auto_filter for data tables
  - Add charts to visualize key metrics
  - Use conditional_formatting to highlight important values
  - Use merge_cells for section headers spanning multiple columns""",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Path to save the .xlsx file (relative to workspace)"},
                    "sheets_data": {
                        "type": "array",
                        "description": "Array of sheet definitions with data, styling, charts, and formatting options as described above.",
                        "items": {"type": "object"}
                    },
                    "workbook_settings": {
                        "type": "object",
                        "description": "Optional workbook-level settings"
                    }
                },
                "required": ["file_path", "sheets_data"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "create_file",
            "description": "Create a new file in the AI workspace",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Path to the file"},
                    "content": {"type": "string", "description": "Content of the file"}
                },
                "required": ["file_path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "read_file",
            "description": "Read a file from the AI workspace",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Path to the file"}
                },
                "required": ["file_path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "generate_pptx_presentation",
            "description": """Create professional, beautifully designed PPTX presentations with Ultra PPT Engine v2.0.
Features: gradient backgrounds, decorative shapes, native charts, card layouts, 8 preset themes, and 12+ slide types.

PRESENTATION SETTINGS (presentation_settings):
  - theme: "ocean"|"forest"|"sunset"|"royal"|"midnight"|"coral"|"tech"|"elegant" (default: "midnight")
  - default_font: font name (default: "微软雅黑")
  - title: document title metadata
  - author: document author metadata
  - custom_colors: {primary, secondary, accent, ...} to override theme colors

SLIDE TYPES (slide_type field in each slide):

1. "title" / "cover" — Title/cover slide with gradient background and decorative elements
   {slide_type:"title", title:"Main Title", subtitle:"Subtitle", author:"Name", date:"2024"}

2. "section" / "divider" — Section divider with large text
   {slide_type:"section", title:"Section Name", subtitle:"Description", section_number:1}

3. "content" / "text" — Standard content slide with title bar and body
   {slide_type:"content", title:"Slide Title", content:{text:"Body text"} or content:{bullets:["item1","item2"]} or content:{paragraphs:[{text:"p1",bold:true},{text:"p2"}]}}
   Optional: elements:[{type:"chart",...},{type:"image",...}] for additional elements

4. "two_column" / "two_columns" — Two-column layout with cards
   {slide_type:"two_column", title:"Title", left_title:"Left", right_title:"Right", left_column:{text:"..."}, right_column:{bullets:["a","b"]}}

5. "three_column" / "three_columns" — Three-column card layout
   {slide_type:"three_column", title:"Title", columns:[{icon:"01", title:"Col1", content:"text"}, {icon:"02", title:"Col2", content:{bullets:["a"]}}, ...]}

6. "cards" — Grid of 2-6 cards
   {slide_type:"cards", title:"Title", cards:[{icon:"▶", title:"Card1", content:"desc"}, ...]}

7. "chart" — Native PowerPoint chart slide
   {slide_type:"chart", title:"Title", chart:{chart_type:"column"|"bar"|"line"|"pie"|"area"|"doughnut"|"scatter", categories:["A","B"], series:[{name:"S1",values:[10,20]}], title:"Chart Title", colors:["#hex"], show_legend:true, show_data_labels:true}, description:"Optional side text"}

8. "stats" / "statistics" — Key statistics/numbers showcase
   {slide_type:"stats", title:"Title", stats:[{value:"98%", label:"Accuracy", icon:"✓", description:"extra info"}, ...]}

9. "timeline" — Horizontal timeline
   {slide_type:"timeline", title:"Title", steps:[{title:"Step1", description:"desc", time_label:"Q1 2024"}, ...]}

10. "table" — Professional styled table
    {slide_type:"table", title:"Title", table:{data:[["H1","H2"],["v1","v2"]], header_bg:"#hex", header_fg:"#fff", font_size:12}}

11. "image" — Image with optional description
    {slide_type:"image", title:"Title", image_path:"path/to/img.png", caption:"Figure 1", description:"Side text"}

12. "comparison" — Side-by-side comparison (Pros/Cons, Before/After)
    {slide_type:"comparison", title:"Title", left_title:"Pros", right_title:"Cons", left_items:["item1"], right_items:["item2"]}

13. "quote" — Large centered quote with author
    {slide_type:"quote", quote:"Quote text here", author:"Author Name"}

14. "ending" / "thank_you" — Thank-you/ending slide
    {slide_type:"ending", title:"Thank You!", subtitle:"Questions?", contact:"email@example.com"}

TEXT CONTENT FORMATS (for content, left_column, right_column, etc.):
  - Simple text: {text: "Hello world"}
  - Bullet list: {bullets: ["item1", "item2", {text:"item3", bold:true}]}
  - Rich paragraphs: {paragraphs: [{text:"p1", bold:true, color:"#hex"}, {text:"p2", font_size:14}]}
  - Common options: font_size, color, bold, italic, align("LEFT"|"CENTER"|"RIGHT"), line_spacing

NATIVE CHART TYPES: column, column_stacked, bar, bar_stacked, line, line_smooth, pie, doughnut, area, area_stacked, scatter

TIPS FOR BEAUTIFUL PRESENTATIONS:
  - Always start with a "title" slide and end with an "ending" slide
  - Use "section" slides to divide major parts
  - Use "stats" for key numbers, "chart" for data visualization
  - Use "cards" or "three_column" for feature/benefit lists
  - Use "comparison" for pros/cons or before/after
  - Use "timeline" for processes, roadmaps, or history
  - Use "quote" for impactful statements
  - Choose a theme that matches the topic (e.g. "tech" for technology, "elegant" for business)
  - All slides automatically get gradient/solid backgrounds, decorative elements, and page numbers""",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Path to save the .pptx file (relative to workspace)"},
                    "slides_content": {
                        "type": "array",
                        "items": {"type": "object"},
                        "description": "List of slide definitions. Each slide MUST have a 'slide_type' field."
                    },
                    "presentation_settings": {
                        "type": "object",
                        "description": "Global settings: theme, default_font, title, author, custom_colors"
                    }
                },
                "required": ["file_path", "slides_content"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "list_directory",
            "description": "List directory contents in the AI workspace",
            "parameters": {
                "type": "object",
                "properties": {
                    "dir_path": {"type": "string", "description": "Path to the directory", "default": "."}
                }
            }
        }
    }
]
