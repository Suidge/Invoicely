from PIL import Image, ImageDraw, ImageFilter
import math

def create_icon():
    size = 1024
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Background Gradient
    for y in range(size):
        r = int(30 + (y / size) * 10)
        g = int(41 + (y / size) * 10)
        b = int(59 + (y / size) * 10)
        draw.line([(0, y), (size, y)], fill=(r, g, b))

    # Paper Shadow
    shadow_offset = 20
    shadow_blur = 40
    # Create a separate image for shadow to blur it
    shadow_img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow_img)
    
    paper_x, paper_y = 212, 162
    paper_w, paper_h = 600, 700
    radius = 60
    
    shadow_draw.rounded_rectangle(
        [(paper_x + shadow_offset, paper_y + shadow_offset), 
         (paper_x + paper_w + shadow_offset, paper_y + paper_h + shadow_offset)],
        radius=radius,
        fill=(0, 0, 0, 80)
    )
    shadow_img = shadow_img.filter(ImageFilter.GaussianBlur(radius=shadow_blur))
    img.paste(shadow_img, (0, 0), shadow_img)

    # Paper Base
    draw.rounded_rectangle(
        [(paper_x, paper_y), (paper_x + paper_w, paper_y + paper_h)],
        radius=radius,
        fill=(248, 250, 252)
    )

    # Folded Corner
    fold_size = 80
    draw.polygon(
        [(paper_x + paper_w - fold_size, paper_y),
         (paper_x + paper_w, paper_y + fold_size),
         (paper_x + paper_w, paper_y)],
        fill=(226, 232, 240)
    )
    # Fold line
    draw.line(
        [(paper_x + paper_w - fold_size, paper_y),
         (paper_x + paper_w, paper_y + fold_size)],
        fill=(203, 213, 225),
        width=2
    )

    # Content Lines
    line_color = (203, 213, 225)
    line_y_start = 300
    line_spacing = 60
    
    # Header Line (Blue)
    draw.rounded_rectangle(
        [(paper_x + 60, line_y_start), (paper_x + 300, line_y_start + 24)],
        radius=12,
        fill=(59, 130, 246)
    )
    
    # Subtitle Line
    draw.rounded_rectangle(
        [(paper_x + 60, line_y_start + 80), (paper_x + 200, line_y_start + 80 + 16)],
        radius=8,
        fill=line_color
    )

    # Body Lines
    for i in range(3):
        y = line_y_start + 160 + (i * line_spacing)
        width = 480 - (i * 40)
        draw.rounded_rectangle(
            [(paper_x + 60, y), (paper_x + 60 + width, y + 16)],
            radius=8,
            fill=line_color
        )

    # Accent Circle (Green)
    circle_x = paper_x + 480
    circle_y = line_y_start + 200
    circle_r = 60
    
    # Circle Shadow
    circle_shadow = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    cs_draw = ImageDraw.Draw(circle_shadow)
    cs_draw.ellipse(
        [(circle_x - circle_r + 5, circle_y - circle_r + 5),
         (circle_x + circle_r + 5, circle_y + circle_r + 5)],
        fill=(0, 0, 0, 40)
    )
    circle_shadow = circle_shadow.filter(ImageFilter.GaussianBlur(radius=10))
    img.paste(circle_shadow, (0, 0), circle_shadow)

    # Circle Base
    draw.ellipse(
        [(circle_x - circle_r, circle_y - circle_r),
         (circle_x + circle_r, circle_y + circle_r)],
        fill=(16, 185, 129)
    )

    # Checkmark
    check_x = circle_x
    check_y = circle_y
    # Simple checkmark path
    check_points = [
        (check_x - 25, check_y),
        (check_x - 5, check_y + 20),
        (check_x + 30, check_y - 25)
    ]
    draw.line(check_points, fill=(255, 255, 255), width=12)

    # Save
    img.save('assets/app_icon.png')
    print("Icon generated successfully!")

if __name__ == '__main__':
    create_icon()
