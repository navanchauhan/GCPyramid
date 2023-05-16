import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename, asksaveasfilename

class CompanySelector:
    def __init__(self, master):
        self.master = master
        self.master.title("Company Selector")

        self.select_file_button = tk.Button(self.master, text="Select Excel File", command=self.load_file)
        self.select_file_button.pack()

    def load_file(self):
        file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        self.df = pd.read_excel(file_path)
        self.show_companies()

    def show_companies(self):
        self.select_file_button.pack_forget()

        self.company_listbox = tk.Listbox(self.master, selectmode=tk.MULTIPLE, exportselection=False)
        self.company_listbox.pack()

        for company in self.df["Symbol"]:
            self.company_listbox.insert(tk.END, company)

        self.create_pyramid_button = tk.Button(self.master, text="Create Pyramid", command=self.create_pyramid)
        self.create_pyramid_button.pack()

    def create_pyramid(self):
        selected_companies = [self.company_listbox.get(index) for index in self.company_listbox.curselection()]
        selected_df = self.df[self.df["Symbol"].isin(selected_companies)].sort_values(by="Weighting", ascending=False)
        #pyramid_file_path = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")])



        image_file_path = asksaveasfilename(defaultextension=".png", filetypes=[("Image files","*.png")])

        company_scores = {}

        # For each 'Company Name', company_scores[company] = 'Weighting' for that company
        for company, ticker, weighting in zip(selected_df["Company Name"], selected_df["Symbol"], selected_df["Weighting"]):
            company_scores[f'{company} ({ticker})'] = weighting

        from PIL import Image, ImageDraw, ImageFont
        import math
        # Sort the dictionary by score in ascending order
        company_scores = {k: v for k, v in sorted(company_scores.items(), key=lambda item: item[1])}

        # Group the companies by their scores
        pyramid = {i: [] for i in range(1, 7)}
        for company, score in company_scores.items():
            pyramid[score].append(company)

        # Initialize some parameters
        img_width = 3300
        img_height = 2550
        block_color = (0, 0, 255)  # Blue color
        font_color = (255, 255, 255)  # White color
        padding = 10  # Padding around blocks
        radius = 20 
        max_companies_in_row = 5

        # Group the companies by their scores
        pyramid = {}
        for company, score in company_scores.items():
            if score not in pyramid:
                pyramid[score] = []
            pyramid[score].append(company)

        # Sort the pyramid keys in ascending order to have a pyramid shape
        sorted_keys = sorted(pyramid.keys())

        companies = []

        for key in sorted_keys:
            if pyramid[key] == []:
                continue
            if len(pyramid[key])<=max_companies_in_row:
                companies.append(pyramid[key])
            else:
                temp_list = []
                for x in range(len(pyramid[key])//max_companies_in_row):
                    temp_list.append(pyramid[key][x*max_companies_in_row:max_companies_in_row])

                temp_list.append(pyramid[key][::-1][:len(pyramid[key])%max_companies_in_row])
                for row in temp_list[::-1]:
                    if row != []:
                        companies.append(row)

        for company_row in companies:
            print(company_row)


        # Calculate the maximum number of companies in a group (this will be the width of your pyramid)
        max_companies = max_companies_in_row #max(len(row) for row in companies) #max(len(v) for v in pyramid.values())

        # Calculate the total number of groups (the height of your pyramid)
        total_groups = len(companies) #len(pyramid)

        # Calculate the size of each block based on the width and height of the image and the number of blocks
        block_size = min((img_width - padding) // max_companies - padding, (img_height - padding) // total_groups - padding)

        # Calculate the size of each block based on the width and height of the image and the number of blocks
        block_width = (img_width - padding) // max_companies - padding
        block_height = (img_height - padding) // total_groups - padding

        # Calculate the total width and height of the blocks (including padding)
        total_width = max_companies * (block_width + padding)
        total_height = total_groups * (block_height + padding)

        # Calculate the starting position for the first block
        start_x = (img_width - total_width) // 2
        start_y = (img_height - total_height) // 2


        # Create an image big enough to hold the pyramid
        img = Image.new('RGB', (img_width, img_height), "white")
        d = ImageDraw.Draw(img)

        def wrap_text(text, max_length):
            words = text.split()
            lines = []
            current_line = []

            for word in words:
                if len(' '.join(current_line + [word])) <= max_length:
                    current_line.append(word)
                else:
                    lines.append(' '.join(current_line))
                    current_line = [word]
            lines.append(' '.join(current_line))
            
            return '\n'.join(lines)

        # Loop over each level of the pyramid
        for i, row in enumerate(companies):
            for j, company in enumerate(row):
                # Calculate the position of the block
                x = start_x + j * (block_width + padding) + (max_companies - len(row)) * (block_width + padding) // 2
                y = start_y + i * (block_size + padding)

                # Draw the block
                d.rounded_rectangle([x, y, x + block_width, y + block_height], fill=block_color, radius=radius)

                # Adjust font size based on the length of the company name and block size
                font_size = min(block_width // (len(company) // 2 + 1), block_height // 2)
                fnt = ImageFont.truetype('./assets/arial.ttf', font_size)

                # Implement word wrap for the company name
                wrapped_company = wrap_text(company, block_width // font_size)

                # Draw the company name
                bbox = d.textbbox((x, y), wrapped_company, font=fnt)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                text_x = x + (block_width - text_width) // 2
                text_y = y + (block_height - text_height) // 2
                d.text((text_x, text_y), wrapped_company, font=fnt, fill=font_color)


        print(image_file_path)

        # Save the image
        if not image_file_path:
            return

        img.save(image_file_path)

        """if not pyramid_file_path:
                                    return
                        
                                with pd.ExcelWriter(pyramid_file_path) as writer:
                                    selected_df.to_excel(writer, index=False)"""

        self.master.quit()

if __name__ == "__main__":
    root = tk.Tk()
    app = CompanySelector(root)
    root.mainloop()
