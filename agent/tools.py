"""
Tool definitions for Main AI Chat and PPT AI Agent
OpenAI-compatible format for use with OpenRouter
"""

# ===== MAIN AI CHAT TOOLS =====

# Generate PPT - Custom tool
GENERATE_PPT_TOOL = {
    "type": "function",
    "function": {
        "name": "generate_ppt",
        "description": """
        This tool calls an AI agent that will generate a complete presentation as a series of html slides and files.
        Use this tool ONLY when you have collected ALL required information from the user:
        - Topic of the presentation
        - Description and purpose
        - Detailed content outline or key points
        - Brand colors
        - Logo details
        - Brand guidelines
        - Additional data or statistics

        Please ensure that you are passing as much relevant and detailed information you can to the AI agent to create the best possible presentation.
        """,
        "parameters": {
            "type": "object",
            "properties": {
                "ppt_topic": {
                    "type": "string",
                    "description": "The main topic/title of the presentation (e.g., 'Q4 Product Roadmap 2025')"
                },
                "ppt_description": {
                    "type": "string",
                    "description": "A detailed description of what the presentation is about and its purpose"
                },
                "ppt_details": {
                    "type": "string",
                    "description": "Detailed content outline, key points to cover, data to include, and overall structure"
                },
                "ppt_data": {
                    "type": "string",
                    "description": "Any specific data, statistics, or numbers to include. Any logo asset file links which can be passed to the ppt agent to use can also be passed here."
                },
                "brand_logo_details": {
                    "type": "string",
                    "description": "Details about the brand logo - file path, URL, or description"
                },
                "brand_guideline_details": {
                    "type": "string",
                    "description": "Brand guidelines including tone, voice, style preferences, font types anything relevant"
                },
                "brand_color_details": {
                    "type": "string",
                    "description": "Brand colors in hex format (e.g., 'primary: #1E40AF, secondary: #F59E0B')"
                }
            },
            "required": ["ppt_topic", "ppt_description", "ppt_details", "ppt_data", "brand_logo_details", "brand_guideline_details", "brand_color_details"]
        }
    }
}

# Web Search tool (OpenAI-compatible format)
WEB_SEARCH_TOOL = {
    "type": "function",
    "function": {
        "name": "web_search",
        "description": "Search the web for current information. Use this when you need up-to-date data, brand guidelines, statistics, or any information that might not be in your training data.",
        "parameters": {
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "The search query"
                }
            },
            "required": ["query"]
        }
    }
}


# ===== PPT AI AGENT TOOLS =====

CREATE_FOLDER_TOOL = {
    "type": "function",
    "function": {
        "name": "create_folder",
        "description": """Create a new folder in the slides directory.
        Use this to organize slide files or assets.
        """,
        "parameters": {
            "type": "object",
            "properties": {
                "folder_path": {
                    "type": "string",
                    "description": "Path of the folder to create (relative to project root, e.g., 'slides/assets')"
                }
            },
            "required": ["folder_path"]
        }
    }
}

CREATE_FILE_TOOL = {
    "type": "function",
    "function": {
        "name": "create_file",
        "description": """Create a new HTML slide file with complete, valid HTML content or a css file with complete valid css content, for css always use tailwind css import from cdn in html files when needed, except the custom base css.

        The content parameter should contain the COMPLETE HTML file, not just a snippet.
        """,
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path where the file should be created (e.g., 'slides/slide_1.html')"
                },
                "content": {
                    "type": "string",
                    "description": "Complete HTML file content including DOCTYPE, head, and body, cdn imports of tailwind css, google fonts, icon packs etc when needed, or complete valid css content for css files"
                }
            },
            "required": ["file_path", "content"]
        }
    }
}

READ_FILE_TOOL = {
    "type": "function",
    "function": {
        "name": "read_file",
        "description": """Read the contents of an existing file.
        Use this to review previously created slides or configuration files.
        """,
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path of the file to read"
                }
            },
            "required": ["file_path"]
        }
    }
}

UPDATE_FILE_TOOL = {
    "type": "function",
    "function": {
        "name": "update_file",
        "description": """Update an existing HTML slide file with corrected or modified content.
        The content parameter should be the COMPLETE updated HTML file.
        """,
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path of the file to update"
                },
                "content": {
                    "type": "string",
                    "description": "Complete updated HTML file content"
                }
            },
            "required": ["file_path", "content"]
        }
    }
}

LIST_FILES_TOOL = {
    "type": "function",
    "function": {
        "name": "list_files",
        "description": """List all files in a directory.
        Use this to see what slides have been created.
        """,
        "parameters": {
            "type": "object",
            "properties": {
                "directory": {
                    "type": "string",
                    "description": "Directory to list files from (default: 'slides')"
                }
            },
            "required": []
        }
    }
}

RETURN_PPT_RESULT_TOOL = {
    "type": "function",
    "function": {
        "name": "return_ppt_result",
        "description": """Return the final result of the PPT generation process.
        Use this ONLY when all slides have been created and you're ready to finish.

        This tool marks the completion of the PPT generation process.
        """,
        "parameters": {
            "type": "object",
            "properties": {
                "success": {
                    "type": "boolean",
                    "description": "Whether the PPT generation was successful"
                },
                "message": {
                    "type": "string",
                    "description": "Summary message about the generated presentation"
                },
                "slide_count": {
                    "type": "integer",
                    "description": "Number of slides created"
                },
                "slide_files": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of slide file paths created"
                }
            },
            "required": ["success", "message", "slide_count", "slide_files"]
        }
    }
}


# ===== NATIVE PPT TOOLS (editable PPTX) =====

CREATE_SLIDE_JSON_TOOL = {
    "type": "function",
    "function": {
        "name": "create_slide_json",
        "description": """Create a single presentation slide by providing structured JSON.
        Each slide has a background and a list of elements (text boxes, shapes, tables).

        COORDINATE SYSTEM: Slide is 10 inches wide x 5.625 inches tall. Origin is top-left corner.
        All position values (left, top, width, height) are in INCHES.

        ELEMENT TYPES:
        1. "text_box" - A text box with formatted text
           - content: string or array of strings/objects for bullet points
           - position: {left, top, width, height} in inches
           - style: {font_name, font_size (pt), font_bold, font_italic, font_color (#hex), alignment (left/center/right), background_color (#hex)}

        2. "shape" - A geometric shape (rectangle, rounded_rectangle, oval, circle, triangle, diamond, chevron, arrow_right)
           - shape_type: one of the above types
           - position: {left, top, width, height} in inches
           - style: {fill_color (#hex), line_color (#hex), line_width (pt)}
           - text: optional text inside the shape
           - style.text_style: optional font styling for shape text

        3. "table" - A data table
           - data: 2D array of strings, first row is header
           - position: {left, top, width, height} in inches
           - style: {header_color (#hex), header_font_color (#hex), font_size (pt)}

        SAFE MARGINS: Keep content within 0.8" from all edges. Usable area: left 0.8 to 9.2, top 0.8 to 4.825.

        DESIGN TIPS:
        - Title text: font_size 36-44pt, font_bold true
        - Subtitle: font_size 20-24pt
        - Body text: font_size 16-18pt
        - Use shapes as decorative accents (colored bars, backgrounds)
        - For bullet lists, use content as array: ["Point 1", "Point 2", "Point 3"]
        """,
        "parameters": {
            "type": "object",
            "properties": {
                "slide_number": {
                    "type": "integer",
                    "description": "Sequential slide number (1, 2, 3...)"
                },
                "background": {
                    "type": "object",
                    "description": "Slide background. Use {\"color\": \"#HEXCODE\"} for solid color, or omit for white.",
                    "properties": {
                        "color": {"type": "string", "description": "Background color in hex (e.g. '#1E40AF')"}
                    }
                },
                "elements": {
                    "type": "array",
                    "description": "List of slide elements (text_box, shape, table)",
                    "items": {
                        "type": "object",
                        "properties": {
                            "type": {"type": "string", "description": "Element type: text_box, shape, table, image"},
                            "content": {"description": "Text content (string or array for bullets)"},
                            "text": {"type": "string", "description": "Text inside a shape"},
                            "shape_type": {"type": "string", "description": "Shape type for shape elements"},
                            "data": {"type": "array", "description": "2D array for table data"},
                            "image_path": {"type": "string", "description": "Path to image file"},
                            "position": {
                                "type": "object",
                                "properties": {
                                    "left": {"type": "number"},
                                    "top": {"type": "number"},
                                    "width": {"type": "number"},
                                    "height": {"type": "number"}
                                },
                                "required": ["left", "top", "width", "height"]
                            },
                            "style": {"type": "object", "description": "Styling properties"}
                        },
                        "required": ["type", "position"]
                    }
                }
            },
            "required": ["slide_number", "elements"]
        }
    }
}

RETURN_NATIVE_PPT_RESULT_TOOL = {
    "type": "function",
    "function": {
        "name": "return_ppt_result",
        "description": """Return the final result of the PPT generation process.
        Use this ONLY when all slides have been created via create_slide_json and you're ready to finish.
        Set slide_files to an empty array since native mode does not create HTML files.
        """,
        "parameters": {
            "type": "object",
            "properties": {
                "success": {
                    "type": "boolean",
                    "description": "Whether the PPT generation was successful"
                },
                "message": {
                    "type": "string",
                    "description": "Summary message about the generated presentation"
                },
                "slide_count": {
                    "type": "integer",
                    "description": "Number of slides created"
                },
                "slide_files": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Set to empty array [] for native mode"
                }
            },
            "required": ["success", "message", "slide_count", "slide_files"]
        }
    }
}


# Collect all tools for easy access
MAIN_CHAT_TOOLS = [
    WEB_SEARCH_TOOL,
    GENERATE_PPT_TOOL
]

# HTML/screenshot mode tools (original)
PPT_AGENT_TOOLS = [
    CREATE_FOLDER_TOOL,
    CREATE_FILE_TOOL,
    READ_FILE_TOOL,
    UPDATE_FILE_TOOL,
    LIST_FILES_TOOL,
    RETURN_PPT_RESULT_TOOL
]

# Native editable PPTX mode tools
PPT_AGENT_NATIVE_TOOLS = [
    CREATE_SLIDE_JSON_TOOL,
    RETURN_NATIVE_PPT_RESULT_TOOL
]
