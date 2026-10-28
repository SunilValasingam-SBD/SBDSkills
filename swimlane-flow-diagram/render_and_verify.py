"""
Self-verification helper for swimlane pptx flow diagrams.

Opens the given .pptx with PowerPoint (COM automation), forces every connector to
reroute cleanly (RerouteConnections), saves the file, then exports each slide as a
high-resolution PNG so you can visually inspect the result before delivering it.

Usage:
    py render_and_verify.py "C:\\path\\to\\Flow_Diagram.pptx"

Requires: pywin32 (py -m pip install pywin32) and a local install of PowerPoint.
"""

import sys
import os
import win32com.client


def render(pptx_path):
    out_dir = os.path.dirname(os.path.abspath(pptx_path))
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1
    pres = powerpoint.Presentations.Open(pptx_path, WithWindow=False)

    msoConnector = 2
    for slide in pres.Slides:
        for shape in slide.Shapes:
            if shape.Type == msoConnector:
                try:
                    shape.RerouteConnections()
                except Exception as e:
                    print("reroute failed:", e)

    pres.Save()

    previews = []
    for i, slide in enumerate(pres.Slides, start=1):
        out_png = os.path.join(out_dir, f"slide_{i}.png")
        slide.Export(out_png, "PNG", 2400, 1350)
        previews.append(out_png)
        print("Exported:", out_png)

    pres.Close()
    powerpoint.Quit()
    return previews


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: py render_and_verify.py <path-to-pptx>")
        sys.exit(1)
    render(sys.argv[1])
