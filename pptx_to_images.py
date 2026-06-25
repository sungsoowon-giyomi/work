"""
PowerPoint COM 자동화로 슬라이드를 PNG 이미지로 변환
"""
import sys
import os
from pathlib import Path
import win32com.client
import pythoncom

def pptx_to_images(pptx_path, out_dir, dpi=150):
    pptx_path = str(Path(pptx_path).resolve())
    out_dir = Path(out_dir)
    out_dir.mkdir(exist_ok=True)

    pythoncom.CoInitialize()
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = True  # Windows requires visible=True for some operations

    try:
        presentation = ppt.Presentations.Open(pptx_path, ReadOnly=True, Untitled=False, WithWindow=False)
        n_slides = presentation.Slides.Count
        print(f"슬라이드 수: {n_slides}")

        # 해상도 설정 (width in pixels)
        # 16:9 at 150dpi ≈ 2400x1350
        width_px = int(16 * dpi)
        height_px = int(9 * dpi)

        saved = []
        for i in range(1, n_slides + 1):
            slide = presentation.Slides(i)
            out_path = out_dir / f"slide-{i:02d}.png"
            slide.Export(str(out_path.resolve()), "PNG", width_px, height_px)
            print(f"  [{i}/{n_slides}] {out_path.name}")
            saved.append(str(out_path))

        presentation.Close()
        return saved
    finally:
        ppt.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    pptx = sys.argv[1]
    out = sys.argv[2]
    max_slides = int(sys.argv[3]) if len(sys.argv) > 3 else 999

    saved = pptx_to_images(pptx, out)
    print(f"\n완료: {len(saved[:max_slides])}개 이미지")
