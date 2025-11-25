# windows_mcp_server.py — OFFICIAL MCP SCHEMA IMPLEMENTATION
# Dependencies Required:
# pip install pyautogui keyboard pywinauto psutil flask pillow

import json
import time
import queue
from flask import Flask, Response, request, stream_with_context
import pyautogui
import keyboard
import psutil
from PIL import ImageGrab
from pywinauto import Application

app = Flask(__name__)
SERVER_PORT = 8000

# Queue to hold incoming MCP 'request' messages from the client stream
request_queue = queue.Queue()


# --- Tool Implementations (Core Windows Automation Functions) ---

def tool_move_mouse(a):
    pyautogui.moveTo(a['x'], a['y'])
    return f"Moved mouse pointer to ({a['x']},{a['y']})"


def tool_click(a):
    pyautogui.click(button=a.get('button', 'left'))
    return f"Performed {a.get('button', 'left')} click"


def tool_type_text(a):
    pyautogui.typewrite(a['text'])
    return f"Typed text: {a['text']}"


def tool_press_key(a):
    keyboard.press_and_release(a['key'])
    return f"Pressed key(s): {a['key']}"


def tool_screenshot(a):
    ImageGrab.grab().save('mcp_screenshot.png')
    return "Screenshot saved to mcp_screenshot.png"


def tool_open_app(a):
    Application().start(a['path'])
    return f"Application started from path: {a['path']}"


def tool_close_app(a):
    psutil.Process(a['pid']).terminate()
    return f"Process with PID {a['pid']} terminated"


def tool_find_window(a):
    try:
        Application().connect(title_re=a['title'])
        return "True"
    except:
        return "False"


def tool_focus_window(a):
    try:
        app = Application().connect(title_re=a['title'])
        app.top_window().set_focus()
        return f"Window '{a['title']}' focused."
    except Exception as e:
        return f"Error focusing window: {e}"


def tool_resize_window(a):
    try:
        app = Application().connect(title_re=a['title'])
        app.top_window().resize(a['width'], a['height'])
        return f"Window '{a['title']}' resized to {a['width']}x{a['height']}"
    except Exception as e:
        return f"Error resizing window: {e}"


def tool_pixel_color(a):
    c = pyautogui.screenshot().getpixel((a['x'], a['y']))
    return json.dumps({"r": c[0], "g": c[1], "b": c[2]})


TOOLS = {
    "move_mouse": tool_move_mouse,
    "click": tool_click,
    "type_text": tool_type_text,
    "press_key": tool_press_key,
    "screenshot": tool_screenshot,
    "open_app": tool_open_app,
    "close_app": tool_close_app,
    "find_window": tool_find_window,
    "focus_window": tool_focus_window,
    "resize_window": tool_resize_window,
    "pixel_color": tool_pixel_color,
}

# --- HELPER FUNCTION FOR MANIFEST ---
def get_mcp_manifest():
    return {
        "tools": [
            # MOUSE & SCREEN
            {
                "name": "move_mouse",
                "description": "Move the mouse pointer to screen coordinates (X, Y).",
                "inputSchema": {
                    "type": "object",
                    "properties": {"x": {"type": "number"}, "y": {"type": "number"}},
                    "required": ["x", "y"]
                }
            },
            {
                "name": "click",
                "description": "Perform a mouse click at the current position.",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "button": {"type": "string", "enum": ["left", "right", "middle"]}
                    }
                }
            },
            {
                "name": "pixel_color",
                "description": "Retrieve the RGB color of a pixel at coordinates (X, Y).",
                "inputSchema": {
                    "type": "object",
                    "properties": {"x": {"type": "number"}, "y": {"type": "number"}},
                    "required": ["x", "y"]
                }
            },
            {
                "name": "screenshot",
                "description": "Capture a full screenshot and save it locally.",
                "inputSchema": {"type": "object", "properties": {}}
            },

            # KEYBOARD
            {
                "name": "type_text",
                "description": "Type a string of text.",
                "inputSchema": {
                    "type": "object",
                    "properties": {"text": {"type": "string"}},
                    "required": ["text"]
                }
            },
            {
                "name": "press_key",
                "description": "Press and release a specific key.",
                "inputSchema": {
                    "type": "object",
                    "properties": {"key": {"type": "string"}},
                    "required": ["key"]
                }
            },

            # WINDOW / APPLICATION CONTROL
            {
                "name": "find_window",
                "description": "Check if a window with the given title exists.",
                "inputSchema": {
                    "type": "object",
                    "properties": {"title": {"type": "string"}},
                    "required": ["title"]
                }
            },
            {
                "name": "focus_window",
                "description": "Bring the window with the given title to the foreground.",
                "inputSchema": {
                    "type": "object",
                    "properties": {"title": {"type": "string"}},
                    "required": ["title"]
                }
            },
            {
                "name": "resize_window",
                "description": "Resize a window to new width & height.",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "title": {"type": "string"},
                        "width": {"type": "number"},
                        "height": {"type": "number"}
                    },
                    "required": ["title", "width", "height"]
                }
            },

            # PROCESS CONTROL
            {
                "name": "open_app",
                "description": "Start an application using its full path.",
                "inputSchema": {
                    "type": "object",
                    "properties": {"path": {"type": "string"}},
                    "required": ["path"]
                }
            },
            {
                "name": "close_app",
                "description": "Terminate a process using its PID.",
                "inputSchema": {
                    "type": "object",
                    "properties": {"pid": {"type": "number"}},
                    "required": ["pid"]
                }
            }
        ]
    }


# --------------------------------------------------
# 1. MCP Tools Manifest: (GET /mcp/tools)
# --------------------------------------------------

@app.route('/mcp/tools', methods=['GET'])
def tools_manifest():
    manifest = get_mcp_manifest()
    return Response(json.dumps(manifest), mimetype='application/json')


# --------------------------------------------------
# 2. MCP SSE Stream (POST for requests, GET for responses)
# --------------------------------------------------

@app.route('/mcp/sse', methods=['POST'])
def sse_post():
    try:
        body = request.get_json(silent=True)

        if (body
            and body.get("type") == "request"
            and body.get("method") == "tools/call"
            and "params" in body):
            request_queue.put(body)
            return Response(json.dumps({"status": "received", "id": body.get("id")}),
                            mimetype='application/json')
    except:
        pass

    return Response(json.dumps({"status": "ok"}), mimetype='application/json')


@app.route('/mcp/sse', methods=['GET'])
def sse_get():

    @stream_with_context
    def stream():
        # Immediately send the MCP manifest event when a client connects.
        manifest = get_mcp_manifest()
        yield f"event: manifest\ndata: {json.dumps(manifest)}\n\n"

        # START HEARTBEAT LOOP
        while True:
            # PROCESS INCOMING TOOL CALLS
            if not request_queue.empty():
                req = request_queue.get()
                req_id = req.get("id")
                params = req.get("params", {})
                tool_name = params.get("name")
                args = params.get("arguments", {})

                if tool_name and req_id:
                    try:
                        result = TOOLS[tool_name](args)
                        response = {
                            "type": "response",
                            "id": req_id,
                            "result": {
                                "content": [
                                    {"type": "output_text", "text": str(result)}
                                ]
                            }
                        }
                        yield f"event: message\ndata: {json.dumps(response)}\n\n"

                    except Exception as e:
                        error = {
                            "type": "error",
                            "id": req_id,
                            "error": {
                                "code": 500,
                                "message": f"Tool execution failed: {type(e).__name__}: {str(e)}"
                            }
                        }
                        yield f"event: message\ndata: {json.dumps(error)}\n\n"

            # KEEP-ALIVE
            yield ": heartbeat\n\n"
            time.sleep(1)

    # FIX: Add headers to prevent caching/buffering which breaks Server-Sent Events (SSE).
    return Response(
        stream(),
        mimetype='text/event-stream',
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no"
        }
    )


# --------------------------------------------------
# START SERVER
# --------------------------------------------------

if __name__ == "__main__":
    print("\n--- MCP SERVER ACTIVE (OFFICIAL MCP SCHEMA) ---")
    print(f"URL for 'Add Connector': http://127.0.0.1:{SERVER_PORT}/mcp/sse")
    print("NOTE: Close this window to immediately stop all control.\n")
    app.run(host='127.0.0.1', port=SERVER_PORT, threaded=True)