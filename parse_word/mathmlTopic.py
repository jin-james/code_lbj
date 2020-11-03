from svglib.svglib import svg2rlg
from reportlab.graphics import renderPM
import time


def svgToPng(svg_code):
	t0 = int(round(time.time() * 1000))
	tmp_path = r'C:\Users\j20687\Desktop\%d.svg' % t0
	b = bytes(svg_code, encoding="utf8")
	with open(tmp_path, 'wb') as f:
		f.write(b)
	drawing = svg2rlg(tmp_path)
	png_value = renderPM.drawToString(drawing, fmt="PNG")
	return png_value


if __name__ == '__main__':
	svg_code = """
	    <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#000" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
	        <circle cx="12" cy="12" r="10"/>
	        <line x1="12" y1="8" x2="12" y2="12"/>
	        <line x1="12" y1="16" x2="12" y2="16"/>
	    </svg>
	"""
	png_value = svgToPng(svg_code)
	print(png_value)