"""
PPT AI Agent - Generates slides using AI
Supports two export modes:
  - "native": Creates editable PPTX with real text boxes and shapes (default)
  - "screenshot": Creates HTML slides, screenshots them, inserts as images into PPTX
Uses OpenRouter API (OpenAI-compatible)
"""

from openai import OpenAI
import os
import json
from typing import Dict, List
from .tools import PPT_AGENT_TOOLS, PPT_AGENT_NATIVE_TOOLS
from .tool_executor import PPTToolExecutor
from utils.screenshot import capture_slide_screenshots
from utils.export import create_pptx_from_screenshots
from utils.export_native import create_native_pptx

OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"
MODEL = "z-ai/glm-4.7"


class PPTAgent:
    """AI Agent that generates PowerPoint slides"""

    def __init__(self, api_key: str = None, progress_callback=None, export_mode="native"):
        self.api_key = api_key or os.environ.get("OPENROUTER_API_KEY")
        self.client = OpenAI(
            base_url=OPENROUTER_BASE_URL,
            api_key=self.api_key
        )
        self.tool_executor = PPTToolExecutor()
        self.messages = []
        self.progress_callback = progress_callback
        self.export_mode = export_mode  # "native" or "screenshot"

    def _emit_progress(self, event_type: str, data: dict):
        """Emit a progress event if callback is set"""
        if self.progress_callback:
            self.progress_callback(event_type, data)

    def generate_presentation(self, ppt_data: Dict) -> Dict:
        """Generate a presentation based on the provided data"""

        # Select system prompt and tools based on export mode
        if self.export_mode == "native":
            system_prompt = self._build_native_system_prompt()
            tools = PPT_AGENT_NATIVE_TOOLS
        else:
            system_prompt = self._build_screenshot_system_prompt()
            tools = PPT_AGENT_TOOLS

        user_prompt = self._build_user_prompt(ppt_data)

        self.messages = [
            {"role": "user", "content": user_prompt}
        ]

        print("\n" + "="*60)
        print(f"PPT AI Agent Started (mode: {self.export_mode})")
        print("="*60)
        print(f"\nGenerating presentation: {ppt_data['ppt_topic']}\n")

        self._emit_progress('agent_started', {
            'message': f"PPT AI Agent Started (mode: {self.export_mode})",
            'topic': ppt_data['ppt_topic']
        })

        max_iterations = 30
        iteration = 0
        final_result = None

        while iteration < max_iterations:
            iteration += 1
            print(f"\n--- Iteration {iteration} ---")
            self._emit_progress('iteration', {'iteration': iteration})

            response = self.client.chat.completions.create(
                model=MODEL,
                max_tokens=4000,
                temperature=0,
                messages=[
                    {"role": "system", "content": system_prompt},
                    *self.messages
                ],
                tools=tools,
                tool_choice="auto",
            )

            choice = response.choices[0]

            if choice.message.tool_calls:
                tool_results = self._process_tool_use(choice.message)

                if self._is_generation_complete(tool_results):
                    final_result = self._extract_final_result(tool_results)

                    if final_result.get("success"):
                        final_result = self._export_to_pptx(final_result, ppt_data.get('ppt_topic', 'Presentation'))

                    break

                self.messages.append(choice.message)
                for tool_result in tool_results:
                    self.messages.append(tool_result)

            else:
                print("\nAgent finished without calling return_ppt_result")
                response_text = choice.message.content or ""

                final_result = {
                    "success": False,
                    "message": f"Agent stopped unexpectedly: {response_text}",
                    "slide_count": 0,
                    "slide_files": []
                }
                break

        if iteration >= max_iterations:
            print("\nMax iterations reached")
            final_result = {
                "success": False,
                "message": "Max iterations reached without completion",
                "slide_count": 0,
                "slide_files": []
            }

        print("\n" + "="*60)
        print("PPT AI Agent Finished")
        print("="*60)
        print(f"Result: {final_result}\n")

        return final_result

    def _build_native_system_prompt(self) -> str:
        """System prompt for native editable PPTX mode"""
        return """You are an expert presentation designer. You create slides using the create_slide_json tool, which produces native editable PowerPoint elements (text boxes, shapes, tables).

=== COORDINATE SYSTEM ===
- Slide dimensions: 10 inches wide x 5.625 inches tall (16:9 widescreen)
- Origin: top-left corner (0, 0)
- All position values are in INCHES
- SAFE MARGINS: Keep content within 0.8" from all edges
  - Usable area: left 0.8" to 9.2" (width 8.4"), top 0.8" to 4.825" (height 4.025")

=== ELEMENT TYPES ===

1. **text_box** - For all text content
   - content: string for single line, or array for bullet points ["Point 1", "Point 2"]
   - position: {left, top, width, height} in inches
   - style: {font_name, font_size (in pt), font_bold, font_italic, font_color (#hex), alignment (left/center/right), background_color (#hex)}

2. **shape** - Decorative elements and colored backgrounds
   - shape_type: rectangle, rounded_rectangle, oval, circle, triangle, diamond, chevron, arrow_right
   - position: {left, top, width, height} in inches
   - style: {fill_color (#hex), line_color (#hex or null), line_width (pt)}
   - text: optional text inside the shape
   - style.text_style: {font_name, font_size, font_bold, font_color, alignment} for shape text

3. **table** - Data tables
   - data: 2D array [["Header1", "Header2"], ["Row1Col1", "Row1Col2"]]
   - position: {left, top, width, height} in inches
   - style: {header_color (#hex), header_font_color (#hex), font_size (pt)}

=== DESIGN GUIDELINES ===

**Typography (in points):**
- Title: 36-44pt, bold
- Subtitle: 20-24pt
- Body text: 16-18pt
- Small text: 12-14pt

**Common Layouts:**

Title Slide:
- Colored background shape covering full slide: position {left: 0, top: 0, width: 10, height: 5.625}
- Title text box centered: position {left: 1.5, top: 1.8, width: 7, height: 1.2}
- Subtitle text box: position {left: 1.5, top: 3.2, width: 7, height: 0.8}

Content Slide:
- Title at top: position {left: 0.8, top: 0.5, width: 8.4, height: 0.8}
- Accent bar: shape rectangle, position {left: 0.8, top: 1.3, width: 1.5, height: 0.06}
- Body content: position {left: 0.8, top: 1.6, width: 8.4, height: 3.2}

Two-Column:
- Title: position {left: 0.8, top: 0.5, width: 8.4, height: 0.8}
- Left column: position {left: 0.8, top: 1.6, width: 3.8, height: 3.2}
- Right column: position {left: 5.2, top: 1.6, width: 3.8, height: 3.2}

**Design Principles:**
- Use shapes as colored accent bars, side panels, or background sections
- Keep generous whitespace - do not cram content
- Maximum 5-6 bullet points per slide
- Use brand colors consistently
- Better to have more slides with less content than fewer crowded slides

=== WORKFLOW ===
1. Call create_slide_json for EACH slide (one at a time)
2. After ALL slides are created, call return_ppt_result with slide_files set to []

START creating slides NOW using create_slide_json!
"""

    def _build_screenshot_system_prompt(self) -> str:
        """System prompt for HTML/screenshot mode (original)"""
        return """You are an expert presentation designer who creates HTML slides that will be screenshotted and exported to PowerPoint.

=== CRITICAL TECHNICAL CONSTRAINTS ===

**EXACT SLIDE DIMENSIONS (NON-NEGOTIABLE):**
- Browser viewport: 1920x1080px (screenshots capture this exact viewport)
- Use FULL viewport width and height (100vw x 100vh)
- EVERYTHING must fit within the SAFEBOX AREA - NO SCROLLING, NO OVERFLOW
- Safebox area: viewport minus padding (content must stay within this safe zone to prevent edge cutoff)

**HTML STRUCTURE (REQUIRED FOR EVERY SLIDE):**
```html
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/7.0.1/css/all.min.css">
    <link rel="stylesheet" href="base-styles.css">
</head>
<body class="m-0 p-0 w-screen h-screen overflow-hidden">
    <div class="w-full h-full overflow-hidden flex items-center justify-center p-20">
        <!-- SAFEBOX AREA: All slide content goes here -->
    </div>
</body>
</html>
```

**TAILWIND CSS FIRST (MANDATORY):**
- Use Tailwind utility classes for ALL styling
- ONLY write custom CSS in base-styles.css
- NO custom CSS in individual slide files
- NO inline styles

**PREVENTING CONTENT OVERFLOW (CRITICAL):**
1. Container padding: Use p-16 or p-20 (64-80px) on main content container
2. Maximum content area: ~1760px wide x ~920px tall with p-20 padding
3. Text sizing: Headings max text-6xl, body text text-xl or text-2xl
4. List limits: Maximum 5-6 bullet points per slide
5. Vertical space check: Total height <= 1080px

**DESIGN PRINCIPLES:**
- Think PowerPoint, not website: generous whitespace, limited content per slide
- Better to split into 2 slides than overflow 1 slide

=== WORKFLOW ===
1. Create base-styles.css FIRST
2. Create individual slides
3. Call return_ppt_result

NO animations, transitions, hover effects, or JavaScript.

START WITH base-styles.css NOW!
"""

    def _build_user_prompt(self, ppt_data: Dict) -> str:
        """Build the user prompt with PPT requirements"""
        prompt = f"""Please generate a presentation with the following details:

**Topic**: {ppt_data['ppt_topic']}

**Description**: {ppt_data['ppt_description']}

**Details**: {ppt_data['ppt_details']}

**Data/Statistics**: {ppt_data.get('ppt_data', 'N/A')}
**Brand Colors**: {ppt_data.get('brand_color_details', 'N/A')}
**Logo Details**: {ppt_data.get('brand_logo_details', 'N/A')}
**Brand Guidelines**: {ppt_data.get('brand_guideline_details', 'N/A')}
"""
        return prompt

    def _process_tool_use(self, message) -> List[Dict]:
        """Process tool use requests from the agent"""
        tool_results = []

        for tool_call in message.tool_calls:
            tool_name = tool_call.function.name
            tool_input = json.loads(tool_call.function.arguments)

            print(f"\nTool: {tool_name}")
            print(f"   Input: {str(tool_input)[:300]}")

            self._emit_progress('tool_use', {
                'tool': tool_name,
                'input': tool_input
            })

            result = self.tool_executor.execute_tool(tool_name, tool_input)

            print(f"   Result: {result[:200]}..." if len(result) > 200 else f"   Result: {result}")

            self._emit_progress('tool_result', {
                'tool': tool_name,
                'result': result[:200] if len(result) > 200 else result
            })

            tool_results.append({
                "role": "tool",
                "tool_call_id": tool_call.id,
                "content": result,
            })

        return tool_results

    def _is_generation_complete(self, tool_results: List[Dict]) -> bool:
        """Check if the agent called return_ppt_result"""
        for result in tool_results:
            if "PPT_GENERATION_COMPLETE" in result.get("content", ""):
                return True
        return False

    def _extract_final_result(self, tool_results: List[Dict]) -> Dict:
        """Extract the final result from return_ppt_result"""
        for result in tool_results:
            content = result.get("content", "")
            if "PPT_GENERATION_COMPLETE" in content:
                json_part = content.split("PPT_GENERATION_COMPLETE:")[1].strip()
                return json.loads(json_part)

        return {
            "success": False,
            "message": "Could not extract final result",
            "slide_count": 0,
            "slide_files": []
        }

    def _export_to_pptx(self, result: Dict, presentation_title: str) -> Dict:
        """Export to PPTX - branches based on export mode"""
        try:
            print("\n" + "="*60)
            print(f"Exporting to PPTX (mode: {self.export_mode})")
            print("="*60)

            self._emit_progress('export_started', {
                'message': f'Exporting to PPTX (mode: {self.export_mode})'
            })

            output_file = f"exports/{presentation_title.replace(' ', '_')}.pptx"

            # Native mode: use slides_data JSON
            if self.export_mode == "native" and result.get("slides_data"):
                self._emit_progress('creating_pptx', {
                    'message': 'Creating native editable PPTX...'
                })

                pptx_file = create_native_pptx(
                    slides_data=result["slides_data"],
                    output_file=output_file,
                    presentation_title=presentation_title,
                    progress_callback=self._emit_progress
                )

                result["pptx_file"] = pptx_file
                self._emit_progress('export_complete', {
                    'message': f'Native PPTX created: {pptx_file}'
                })
                return result

            # Screenshot mode: capture HTML slides and insert as images
            slide_files = result.get("slide_files", [])
            if not slide_files:
                print("No slide files to export")
                return result

            self._emit_progress('capturing_screenshots', {
                'message': 'Capturing screenshots...',
                'slide_count': len(slide_files)
            })

            screenshots = capture_slide_screenshots(slide_files, progress_callback=self._emit_progress)

            if not screenshots:
                print("No screenshots captured, skipping PPTX export")
                return result

            self._emit_progress('creating_pptx', {
                'message': 'Creating PPTX presentation...'
            })

            pptx_file = create_pptx_from_screenshots(
                screenshots,
                output_file=output_file,
                presentation_title=presentation_title,
                progress_callback=self._emit_progress
            )

            result["pptx_file"] = pptx_file
            result["screenshots"] = screenshots

            self._emit_progress('export_complete', {
                'message': f'PPTX created: {pptx_file}'
            })

            return result

        except Exception as e:
            print(f"Error exporting to PPTX: {str(e)}")
            import traceback
            traceback.print_exc()
            return result
