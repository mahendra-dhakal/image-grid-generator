# Install necessary libraries
# pip install pandas pillow requests beautifulsoup4

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import zipfile
import shutil
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox
import requests
from bs4 import BeautifulSoup
import urllib.parse
import time
import sys

class ProductGridGenerator:
    def __init__(self):
        # Fixed paths as specified
        self.project_folder = r"C:\Users\mahen\OneDrive\Desktop\image_project"
        self.image_folder = r"C:\Users\mahen\OneDrive\Desktop\image_project\image"
        # FIXED: Output directly to the output folder, no subfolders
        self.output_folder = r"C:\Users\mahen\OneDrive\Desktop\image_project\output"
        self.font_path = None
        
    def setup_directories(self):
        """Create necessary directories and verify project structure"""
        # Create all necessary directories
        os.makedirs(self.image_folder, exist_ok=True)
        os.makedirs(self.output_folder, exist_ok=True)
        
        # Verify project folder exists
        if not os.path.exists(self.project_folder):
            raise Exception(f"Project folder not found: {self.project_folder}")
            
        print(f"✓ Image folder: {self.image_folder}")
        print(f"✓ Output folder: {self.output_folder}")
        
        # Find font file
        font_files = []
        for file in os.listdir(self.project_folder):
            if file.lower().endswith(('.ttf', '.otf')):
                font_files.append(file)
                
        if font_files:
            self.font_path = os.path.join(self.project_folder, font_files[0])
            print(f"✓ Found font: {font_files[0]}")
        else:
            print("⚠ No font file found in project folder. Will use system font.")
            
    def select_csv_file(self):
        """Select CSV file using file dialog"""
        print("\nPlease select your CSV file with columns: Product Name, Quantity, For, Price")
        
        root = tk.Tk()
        root.withdraw()
        
        csv_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialdir=self.project_folder
        )
        
        if not csv_path:
            raise Exception("No CSV file selected!")
            
        return csv_path
        
    def load_fonts(self):
        """Load fonts with optimized sizes for better proportions"""
        try:
            if self.font_path and os.path.exists(self.font_path):
                # Larger font sizes for better visibility (adjusted to fit header)
                header_font = ImageFont.truetype(self.font_path, 120)      # Header
                product_font = ImageFont.truetype(self.font_path, 60)      # Product name
                price_font = ImageFont.truetype(self.font_path, 90)        # Price
                header_attr_font = ImageFont.truetype(self.font_path, 40)  # Sub-header
                print("✓ Loaded custom fonts successfully.")
            else:
                raise IOError("Font not found")
        except IOError:
            print("⚠ Custom font not found. Using system fonts.")
            try:
                # Try to use system Arial font
                header_font = ImageFont.truetype("arial.ttf", 120)
                product_font = ImageFont.truetype("arial.ttf", 60)
                price_font = ImageFont.truetype("arial.ttf", 90)
                header_attr_font = ImageFont.truetype("arial.ttf", 40)
            except:
                # Fall back to default fonts
                header_font = ImageFont.load_default()
                product_font = ImageFont.load_default()
                price_font = ImageFont.load_default()
                header_attr_font = ImageFont.load_default()
        return header_font, product_font, price_font, header_attr_font
        
    def clean_text_for_filename(self, text):
        """Clean text for safe filename usage but keep original for search"""
        text = str(text).replace("–", "-")
        # Remove characters that can't be in filenames
        text = text.replace("/", "_").replace("\\", "_").replace(":", "_")
        text = text.replace("*", "_").replace("?", "_").replace('"', "_")
        text = text.replace("<", "_").replace(">", "_").replace("|", "_")
        return text
        
    def find_local_image(self, product_name):
        """Find image for product in local image folder - search with exact name first"""
        # First try exact product name with common extensions
        extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp']
        
        # Try exact product name first (as downloaded from web)
        for ext in extensions:
            exact_path = os.path.join(self.image_folder, product_name + ext)
            if os.path.exists(exact_path):
                print(f"✓ Found exact match: {product_name}{ext}")
                return exact_path
        
        # If exact match not found, try cleaned versions
        clean_name = self.clean_text_for_filename(product_name).lower()
        
        name_patterns = [
            clean_name,
            clean_name.replace(" ", "_"),
            clean_name.replace(" ", "-"),
            clean_name.replace(" ", ""),
            product_name.lower(),  # Original name
            product_name.lower().replace(" ", "_"),
            product_name.lower().replace(" ", "-"),
            product_name.lower().replace(" ", "")
        ]
        
        for pattern in name_patterns:
            for ext in extensions:
                image_path = os.path.join(self.image_folder, pattern + ext)
                if os.path.exists(image_path):
                    print(f"✓ Found alternative match: {pattern}{ext}")
                    return image_path
                    
        return None
        
    def search_and_download_image(self, product_name):
        """Search for product image on the web and download it with exact product name"""
        try:
            print(f"🔍 Searching web for: {product_name}")
            
            # Search query
            search_query = f"{product_name} product image"
            search_url = f"https://www.google.com/search?q={urllib.parse.quote(search_query)}&tbm=isch"
            
            # Headers to mimic a real browser
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            
            # Get search results
            response = requests.get(search_url, headers=headers, timeout=10)
            if response.status_code != 200:
                print(f"⚠ Search failed for {product_name}")
                return None
                
            # Parse HTML to find image URLs
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Find image elements
            img_elements = soup.find_all('img')
            
            # Try to find actual image URLs from the page
            image_urls = []
            for img in img_elements:
                src = img.get('src')
                if src and src.startswith('http') and any(ext in src.lower() for ext in ['.jpg', '.jpeg', '.png']):
                    image_urls.append(src)
                    
            # Also try to extract from data-src attributes
            for img in img_elements:
                data_src = img.get('data-src')
                if data_src and data_src.startswith('http'):
                    image_urls.append(data_src)
            
            # If no direct URLs found, try alternative approach
            if not image_urls:
                # Look for specific patterns in Google Image search
                scripts = soup.find_all('script')
                for script in scripts:
                    if script.string and 'http' in str(script.string):
                        # Extract URLs from JavaScript
                        import re
                        urls = re.findall(r'https?://[^\s"\']+\.(?:jpg|jpeg|png)', str(script.string))
                        image_urls.extend(urls[:3])  # Take first 3 URLs
                        if image_urls:
                            break
            
            # Try to download the first few images
            for i, url in enumerate(image_urls[:3]):  # Try first 3 URLs
                try:
                    print(f"📥 Attempting to download image {i+1} for {product_name}")
                    
                    img_response = requests.get(url, headers=headers, timeout=15, stream=True)
                    if img_response.status_code == 200:
                        
                        # Verify it's actually an image
                        try:
                            img = Image.open(img_response.raw)
                            
                            # Convert to RGB if necessary
                            if img.mode in ("RGBA", "P"):
                                img = img.convert("RGB")
                                
                            # Resize to standard size for consistency
                            img = img.resize((800, 800))  # Standard size for processing
                            
                            # Save with EXACT product name (no cleaning for filename)
                            image_filename = f"{product_name}.jpg"
                            image_path = os.path.join(self.image_folder, image_filename)
                            
                            img.save(image_path, "JPEG", quality=95)
                            print(f"✓ Downloaded and saved: {image_filename}")
                            
                            return image_path
                            
                        except Exception as e:
                            print(f"⚠ Image processing failed: {e}")
                            continue
                            
                except Exception as e:
                    print(f"⚠ Download failed for URL {i+1}: {e}")
                    continue
                    
            print(f"⚠ Could not download image for {product_name}")
            return None
            
        except Exception as e:
            print(f"⚠ Web search failed for {product_name}: {e}")
            return None
            
    def create_placeholder_image(self, product_name):
        """Create enhanced placeholder image"""
        try:
            placeholder_img = Image.new("RGB", (800, 800), color=(245, 245, 245))
            draw = ImageDraw.Draw(placeholder_img)
            
            # Add border
            draw.rectangle([(0, 0), (799, 799)], outline=(200, 200, 200), width=3)
            
            # Split text to fit in image
            words = product_name.split()
            lines = []
            current_line = ""
            for word in words:
                if len(current_line + word) < 15:
                    current_line += word + " "
                else:
                    lines.append(current_line.strip())
                    current_line = word + " "
            if current_line:
                lines.append(current_line.strip())
                
            # Center the text
            try:
                font = ImageFont.truetype("arial.ttf", 36)
            except:
                font = ImageFont.load_default()
                
            y_start = 400 - (len(lines) * 25)
            for line in lines:
                bbox = draw.textbbox((0, 0), line, font=font)
                text_width = bbox[2] - bbox[0]
                x_center = (800 - text_width) // 2
                draw.text((x_center, y_start), line, fill=(100, 100, 100), font=font)
                y_start += 50
                
            # Add "Image Not Available" at bottom
            draw.text((240, 720), "Image Not Available", fill=(150, 150, 150), font=font)
                
            return placeholder_img
            
        except Exception as e:
            print(f"Error creating placeholder for {product_name}: {e}")
            return None
            
    def process_product_images(self, df):
        """Process all product images - find local, download missing, or create placeholders"""
        print("\n🖼️ Processing Product Images...")
        
        for index, row in df.iterrows():
            product_name = row['Product Name']
            print(f"\nProcessing: {product_name}")
            
            # First, try to find local image with exact name
            local_image_path = self.find_local_image(product_name)
            
            if local_image_path:
                df.at[index, 'Image_Found'] = True
                print(f"✓ Using local image")
                    
            else:
                # Image not found locally, try to download from web
                print(f"Local image not found for: {product_name}")
                
                downloaded_path = self.search_and_download_image(product_name)
                
                if downloaded_path:
                    df.at[index, 'Image_Found'] = True
                    print(f"✓ Downloaded and saved web image")
                        
                else:
                    # Create placeholder and save it
                    print(f"Creating placeholder for: {product_name}")
                    placeholder = self.create_placeholder_image(product_name)
                    
                    if placeholder:
                        placeholder_filename = f"{product_name}.jpg"
                        placeholder_path = os.path.join(self.image_folder, placeholder_filename)
                        placeholder.save(placeholder_path, "JPEG", quality=95)
                        
                        df.at[index, 'Image_Found'] = False
                        print(f"✓ Created and saved placeholder image")
                    else:
                        df.at[index, 'Image_Found'] = False
                        print(f"✗ Failed to create placeholder")
            
            # Small delay to be respectful to web servers
            time.sleep(1)
            
        print(f"\n✓ Completed image processing for {len(df)} products")
        
    def split_text_to_fit(self, draw, text, font, max_width):
        """Split text into multiple lines to fit width"""
        words = text.split()
        lines = []
        current_line = ""
        
        for word in words:
            test_line = current_line + (word + " ")
            line_bbox = draw.textbbox((0, 0), test_line, font=font)
            line_width = line_bbox[2] - line_bbox[0]
            
            if line_width <= max_width:
                current_line = test_line
            else:
                lines.append(current_line.strip())
                current_line = word + " "
                
        if current_line:
            lines.append(current_line.strip())
            
        return lines
        
    def draw_header(self, draw, a4_width, header_font, header_attr_font):
        """Draw page header with proper date formatting"""
        current_date = datetime.now()
        date_from = current_date.strftime("%b %d")
        end_date = current_date + timedelta(days=6)
        date_to = end_date.strftime("%b %d, %Y")
        
        left_text = f"Price effective from {date_from} {current_date.year} to {date_to}"
        right_text = "100 Railroad Avenue, Denmark, WI, United States, Wisconsin"
        
        header_text = "Main Street\nMarket"
        header_text_bbox = draw.textbbox((0, 0), header_text, font=header_font)
        header_text_width = header_text_bbox[2] - header_text_bbox[0]
        
        max_text_width = header_text_width
        
        left_lines = self.split_text_to_fit(draw, left_text, header_attr_font, max_text_width)
        right_lines = self.split_text_to_fit(draw, right_text, header_attr_font, max_text_width)
        
        padding_x = 40
        y_offset_left = 140
        y_offset_right = 140
        
        for line in left_lines:
            draw.text((padding_x, y_offset_left), line, font=header_attr_font, fill=(0, 0, 0))
            y_offset_left += 35
            
        for line in right_lines:
            right_text_bbox = draw.textbbox((0, 0), line, font=header_attr_font)
            text_width = right_text_bbox[2] - right_text_bbox[0]
            draw.text((a4_width - text_width - padding_x, y_offset_right), line, font=header_attr_font, fill=(0, 0, 0))
            y_offset_right += 35
            
    def draw_product_cell(self, draw, grid_image, x, y, product, product_font, price_font, cell_width, cell_height):
        """Draw individual product cell with MUCH LARGER images that fill most of the cell"""
        product_name = product["Product Name"]
        quantity = product.get("Quantity")
        
        display_name = product_name
        if pd.notna(quantity):
            display_name += f" ({quantity})"
            
        max_text_width = cell_width - 40
        product_lines = self.split_text_to_fit(draw, display_name, product_font, max_text_width)
        
        # Draw green background for product name - smaller to leave more room for image
        text_y = y + 15
        highlight_padding = 10
        
        # Calculate total height needed for all lines
        total_text_height = 0
        for line in product_lines:
            line_bbox = draw.textbbox((0, 0), line, font=product_font)
            total_text_height += (line_bbox[3] - line_bbox[1]) + 10
        
        # Draw green background rectangle
        bg_width = max_text_width + (highlight_padding * 2)
        draw.rectangle([(x + 20 - highlight_padding, text_y - highlight_padding),
                       (x + 20 + bg_width, text_y + total_text_height + highlight_padding)],
                      fill=(34, 139, 34))  # Forest green color
        
        # Draw product name text in white
        current_y = text_y
        for line in product_lines:
            draw.text((x + 20, current_y), line, font=product_font, fill=(255, 255, 255))
            line_bbox = draw.textbbox((0, 0), line, font=product_font)
            current_y += (line_bbox[3] - line_bbox[1]) + 10
            
        # Load and draw image - MUCH LARGER to fill most of the cell
        image_path = os.path.join(self.image_folder, f"{product_name}.jpg")
        
        if os.path.exists(image_path):
            try:
                img = Image.open(image_path)
                # Calculate available space for image
                available_width = cell_width - 60  # Small margins
                available_height = cell_height - (current_y - y) - 140  # Leave space for price (less than before)
                # Use 95% of available area
                img_size = int(min(available_width, available_height) * 0.95)
                # Ensure minimum size for visibility
                if img_size < 180:
                    img_size = min(180, available_width, available_height)
                img = img.resize((img_size, img_size), Image.Resampling.LANCZOS)
                img_x = x + (cell_width - img_size) // 2
                img_y = current_y + 20
                grid_image.paste(img, (img_x, img_y))
                print(f"✓ Placed LARGE image for {product_name} at size {img_size}x{img_size}")
            except Exception as e:
                print(f"Error loading image for {product_name}: {e}")
                
        # Draw price at bottom right
        price = str(product["Price"])
        for_value = product.get("For", 1)
        
        if pd.notna(for_value) and for_value > 1:
            price = f"{int(for_value)}/{price}"
            
        price_bbox = draw.textbbox((0, 0), price, font=price_font)
        price_text_width = price_bbox[2] - price_bbox[0]
        
        # Position price at bottom right
        price_x = x + cell_width - price_text_width - 40
        price_y = y + cell_height - 120
        draw.text((price_x, price_y), price, font=price_font, fill=(0, 0, 0))
        
        # Draw cell border
        draw.rectangle([(x + 10, y + 10), (x + cell_width - 30, y + cell_height - 30)], 
                      fill=None, outline=(200, 200, 200), width=3)

    def create_grid_page(self, page_number, products, fonts):
        """Create a single grid page with optimized dimensions matching your reference image"""
        header_font, product_font, price_font, header_attr_font = fonts
        
        # Optimized A4 dimensions - more reasonable size but high quality
        a4_width = 3508   # A4 width at 300 DPI
        a4_height = 4961  # A4 height at 300 DPI
        cell_width = a4_width // 3
        cell_height = (a4_height - 400) // 4  # Adjusted for header
        
        grid_image = Image.new("RGB", (a4_width, a4_height), color=(255, 255, 255))
        draw = ImageDraw.Draw(grid_image)
        
        # Draw main header
        header_text = "Main Street\nMarket"
        header_text_bbox = draw.textbbox((0, 0), header_text, font=header_font)
        header_text_width = header_text_bbox[2] - header_text_bbox[0]
        draw.text(((a4_width - header_text_width) // 2, 80), header_text, font=header_font, fill=(0, 0, 0))
        
        # Draw sub-header
        self.draw_header(draw, a4_width, header_font, header_attr_font)
        
        # Draw products in 3x4 grid
        for idx, product in enumerate(products):
            if idx >= 12:
                break
                
            col = idx % 3
            row = idx // 3
            
            x = col * cell_width + 20
            y = row * cell_height + 300  # Adjusted for header
            
            self.draw_product_cell(draw, grid_image, x, y, product, product_font, price_font, cell_width, cell_height)
            
        # Save with high quality
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file_path = os.path.join(self.output_folder, f"product_grid_{current_time}_page_{page_number}.jpg")
        
        # Save with high quality and good DPI
        grid_image.save(output_file_path, "JPEG", quality=95, dpi=(300, 300))
        print(f"✓ Created OPTIMIZED grid page {page_number} -> {output_file_path}")
        
        return output_file_path
        
    def generate_all_grids(self, df):
        """Generate all grid pages and remove duplicates"""
        # Remove duplicate products based on Product Name
        print(f"Original products: {len(df)}")
        df_unique = df.drop_duplicates(subset=['Product Name'], keep='first')
        print(f"After removing duplicates: {len(df_unique)}")
        
        fonts = self.load_fonts()
        
        items_per_page = 12
        output_files = []
        
        for i in range(0, len(df_unique), items_per_page):
            product_batch = df_unique.iloc[i:i + items_per_page].to_dict(orient='records')
            page_number = (i // items_per_page) + 1
            output_file = self.create_grid_page(page_number, product_batch, fonts)
            output_files.append(output_file)
            
        return output_files
        
    def run(self):
        """Main execution function"""
        print("=== OPTIMIZED Product Grid Generator ===")
        print("🚀 Features:")
        print("✓ Optimized resolution output (3508x4961 pixels - A4 at 300 DPI)")
        print("✓ MUCH LARGER product images that fill most of each cell")
        print("✓ Removes duplicate products automatically")
        print("✓ Saves directly to output folder")
        print("✓ Balanced font sizes and layout")
        print("✓ High quality output matching your reference image")
        
        try:
            # Setup
            self.setup_directories()
            
            # Select CSV file
            csv_path = self.select_csv_file()
            
            # Load CSV
            try:
                df = pd.read_csv(csv_path)
                print(f"✓ Loaded {len(df)} products from CSV")
            except Exception as e:
                raise Exception(f"Error reading CSV file: {e}")
            
            # Validate required columns
            required_columns = ['Product Name', 'Price']
            missing_cols = [col for col in required_columns if col not in df.columns]
            if missing_cols:
                raise Exception(f"Missing required columns: {missing_cols}")
            
            # Process all images (local + web download)
            self.process_product_images(df)
            
            # Generate grids
            print(f"\n🎨 Generating OPTIMIZED Product Grids with LARGE images...")
            output_files = self.generate_all_grids(df)
            
            print(f"\n🎉 Process completed successfully!")
            print(f"✓ Generated {len(output_files)} optimized grid pages with LARGE product images")
            print(f"✓ All images stored in: {self.image_folder}")
            print(f"✓ Grid output saved to: {self.output_folder}")
            print(f"✓ Images now fill most of each cell for better visibility")
            
            # Open output folder
            try:
                os.startfile(self.output_folder)
            except:
                print(f"You can find your files at: {self.output_folder}")
                
        except Exception as e:
            print(f"\n❌ Error: {e}")
            input("Press Enter to exit...")
            
        input("\nPress Enter to exit...")

# Run the generator
if __name__ == "__main__":
    generator = ProductGridGenerator()
    generator.run()