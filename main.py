import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename, asksaveasfilename

from ttkwidgets import ScrolledListbox

import sv_ttk

import os

canadian_companies = []

def pyramid_list(lst, sort_canadian=True):
    old_lst = lst

    top = []
    bottom = []

    for idx, company in enumerate(lst):
        if idx==0:
            top.append(company)
        else:
            if company in canadian_companies:
                bottom.append(company)
            else:
                top.append(company)

    lst = top + bottom

    i = 0
    rows = 0
    pyramid = []

    # Figure out the number of rows
    while i < len(lst):
        rows += 1
        i += rows

    # Generate the pyramid
    i = 0
    for r in range(1, rows+1):
        row = []
        for j in range(r):
            if i < len(lst):
                row.append(lst[i])
                i += 1
        pyramid.append(row)

    if len(pyramid) >= 2:
        if len(pyramid[-1]) < len(pyramid[-2]):
            pyramid[-2].extend(pyramid[-1])
            pyramid.pop()

    print("Checking for canadian companies")
    print(canadian_companies)
    if len(pyramid) > 2:
        print("Pyramid has more than two rows, checking to ensure Canadian companies are at the bottom")
        print(len(pyramid[-1]) - len(pyramid[-2]))
        for row in pyramid:
            print(len(row))
            print(row)
        if ((len(pyramid[-1]) - len(pyramid[-2])) > 1):
            print("Last two rows can be balanced if needed")
            for _ in range((len(pyramid[-1]) - len(pyramid[-2]))-1):
                company_to_add = None
                company_idx = 0
                for idx, company in enumerate(pyramid[-1]):
                    print(f'Checking {company}')
                    if company not in canadian_companies:
                        company_to_add = company
                        company_idx = idx
                        print(f'Shifting {company_to_add}')
                if company_to_add is not None:
                    pyramid[-1].pop(company_idx)
                    pyramid[-2].append(company_to_add)
    
    if len(pyramid) > 3:
        if len(pyramid[-1]) < len(pyramid[-2]):
            pyramid[-3].append(pyramid[-2][-1])
            pyramid[-2].pop(-1)

    return pyramid

class CompanySelector:
    def __init__(self, master):
        self.master = master
        self.master.title("Company Selector")
        self.master.geometry("600x600")
        self.select_file_button = ttk.Button(self.master, text="Select Excel File", command=self.load_file)
        self.select_file_button.pack(pady=250)

    def load_file(self):
        file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        self.df = pd.read_excel(file_path)
        self.df = self.df[self.df['Company Name'].notna()]
        self.df = self.df[self.df['Symbol'].notna()]
        self.show_companies()

    def show_companies(self):
        self.select_file_button.pack_forget()
        self.master.geometry("")

        self.company_listbox = ScrolledListbox(self.master, selectmode=tk.MULTIPLE, exportselection=False, height=20)
        self.company_listbox.pack(expand=True)

        for company in self.df["Symbol"]:
            self.company_listbox.listbox.insert(tk.END, company)


        # bind the listbox to an onselect event
        self.company_listbox.listbox.bind('<<ListboxSelect>>', self.onselect)

        self.num_companies_selected_label = ttk.Label(self.master, text="0 companies selected")
        self.num_companies_selected_label.pack()

        self.pyramid_title_string = tk.StringVar()
        self.pyramid_title_string.set("Copyright (2004-2023) by Gentry Capital Corporation")

        self.pyramid_title =  ttk.Entry(self.master, textvariable=self.pyramid_title_string)
        self.pyramid_title.pack()

        self.create_pyramid_button = ttk.Button(self.master, text="Create Pyramid", command=self.create_pyramid)
        self.create_pyramid_button.pack()

    def onselect(self, evt):
        w = evt.widget
        num_selected = len(w.curselection())
        self.num_companies_selected_label.config(text=f"{num_selected} companies selected")

    def create_pyramid(self):
        selected_companies = [self.company_listbox.listbox.get(index) for index in self.company_listbox.listbox.curselection()]
        selected_df = self.df[self.df["Symbol"].isin(selected_companies)].sort_values(by="Weighting", ascending=False)
        #pyramid_file_path = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")])



        image_file_path = asksaveasfilename(defaultextension=".png", filetypes=[("Image files","*.png")])

        company_scores = {}

        # For each 'Company Name', company_scores[company] = 'Weighting' for that company
        for company, ticker, weighting, dividend, currency in zip(selected_df["Company Name"], selected_df["Symbol"], selected_df["Weighting"], selected_df["Dividend?"], selected_df["Currency"]):
            extra_char = ""
            if dividend.strip() == "Y":
                extra_char = "*"
            if currency == "CAD":
                canadian_companies.append(f'{company} ({ticker}){extra_char}')
            company_scores[f'{company} ({ticker}){extra_char}'] = weighting

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
        img_height = 1550 #2550
        overall_height = 2550
        block_color = (3, 37, 126)  # Blue color
        dividend_color = block_color #(1,50,32)  # Green color
        font_color = (255, 255, 255)  # White color
        normal_block = (3, 37, 126) # Dark Blue  
        padding = 10  # Padding around blocks
        radius = 20 
        logo_padding = 100 # top padding for logo
        pyramid_padding_bottom = 350 # bottom padding for pyramid
        max_companies_in_row = 5
        primary_font = "./assets/baskerville.ttf"
        secondary_font = "./assets/gill_sans_bold.ttf"

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

        everything_list = []
        for row in companies:
            if row == []:
                continue
            for company in row:
                everything_list.append(company)

        companies = everything_list

        companies = pyramid_list(companies)

        for company_row in companies:
            print(company_row)

        # Calculate the maximum number of companies in a group (this will be the width of your pyramid)
        max_companies = max(len(row) for row in companies) #max(len(v) for v in pyramid.values())

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

        min_font_size = 100_000_000

        for i, row in enumerate(companies):
            for j, company in enumerate(row):
                font_size = min(block_width // (len(company) // 2 + 1), block_height // 2)
                if font_size < min_font_size:
                    min_font_size = font_size

        num_dividends = 0

        # Loop over each level of the pyramid
        for i, row in enumerate(companies):
            for j, company in enumerate(row):
                # Calculate the position of the block
                x = start_x + j * (block_width + padding) + (max_companies - len(row)) * (block_width + padding) // 2
                y = start_y + i * (block_size + padding)

                # Calculate the color of the block
                if company[-1] == "*":
                    block_color = dividend_color
                    num_dividends += 1
                else:
                    block_color = normal_block

                # Draw the block
                d.rounded_rectangle([x, y, x + block_width, y + block_height], fill=block_color, radius=radius)

                # Adjust font size based on the length of the company name and block size
                font_size = min_font_size #min(block_width // (len(company) // 2 + 1), block_height // 2)
                fnt = ImageFont.truetype(secondary_font, font_size)

                # Implement word wrap for the company name
                wrapped_company = wrap_text(company, (block_width // font_size)*1.5)

                # Draw the company name
                bbox = d.textbbox((x, y), wrapped_company, font=fnt)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                text_x = x + (block_width - text_width) // 2
                text_y = y + (block_height - text_height) // 2
                d.text((text_x, text_y), wrapped_company, font=fnt, fill=font_color, align='center', spacing=10)


        print(image_file_path)

        # Save the image
        if not image_file_path:
            return

        image_logo = Image.open("./assets/GentryCapitalRGB.jpg")
        _, l_height = image_logo.size

        new_img = Image.new("RGB", (img_width, l_height), "white")
        left = (new_img.width - image_logo.width) // 2
        top = (new_img.height - image_logo.height) // 2

        new_img.paste(image_logo, (left, top))

        final_img = Image.new("RGB", (img_width, overall_height), "white")
        final_img.paste(new_img, (0, logo_padding))

        text1 = self.pyramid_title_string.get()
        text2 = f"({num_dividends} Dividend payors - all identified by asterisk)"
        text3 = "FOR INTERNAL USE ONLY"

        draw = ImageDraw.Draw(final_img)
        font_size = 60

        font1 = ImageFont.truetype(primary_font, font_size)
        text_width1 = draw.textlength(text1, font=font1)
        while text_width1 > final_img.width:
            font_size -= 1
            font1 = ImageFont.truetype(primary_font, font_size)
            text_width1 = draw.textlength(text1, font=font1)

        font2 = ImageFont.truetype(primary_font, font_size)
        text_width2 = draw.textlength(text2, font=font2)
        while text_width2 > final_img.width:
            font_size -= 1
            font2 = ImageFont.truetype(primary_font, font_size)
            text_width2 = draw.textlength(text2, font=font2)

        new_font_size = 30
        font3 = ImageFont.truetype(secondary_font, new_font_size)
        _,t,_,b = draw.textbbox((100,100), text3, font=font3)
        while (b-t) < (0.8*pyramid_padding_bottom):
            new_font_size += 1
            font3 = ImageFont.truetype(secondary_font, new_font_size)
            _,t,_,b = draw.textbbox((100,100), text3, font=font3)
        text_width3 = draw.textlength(text3, font=font3)
        while text_width3 > (final_img.width*0.8):
            new_font_size -= 1
            font3 = ImageFont.truetype(secondary_font, new_font_size)
            text_width3 = draw.textlength(text3, font=font3)
        _,t,_,b = draw.textbbox((100,100),text3,font=font3)

        start_x1 = (final_img.width - text_width1) // 2
        start_y1 = image_logo.height + logo_padding + 50

        start_x2 = (final_img.width - text_width2) // 2
        start_y2 = image_logo.height + logo_padding + 125

        start_x3 = (final_img.width - text_width3) // 2
        start_y3 = (final_img.height - pyramid_padding_bottom + ((pyramid_padding_bottom)-(b-t))//2)
        
        draw.text((start_x1, start_y1), text1, font=font1, fill="black")
        draw.text((start_x2, start_y2), text2, font=font2, fill="black")
        draw.text((start_x3, start_y3), text3, font=font3, fill=(105,105,105))

        final_img.paste(img, (0, l_height + (overall_height - l_height - img_height) - pyramid_padding_bottom) )

        # draw = ImageDraw.Draw(final_img)
        # watermark_text = "FOR INTERNAL USE ONLY"
        # watermark_font = ImageFont.truetype(secondary_font, 72)
        # watermark_width = draw.textlength(watermark_text, watermark_font)
        # draw.text(
        #     ((final_img.width - watermark_width)//2,overall_height-pyramid_padding_bottom),
        #     watermark_text,
        #     font=watermark_font,
        #     align='center'
        #     )

        final_img.save(image_file_path)

        """if not pyramid_file_path:
                                    return
                        
                                with pd.ExcelWriter(pyramid_file_path) as writer:
                                    selected_df.to_excel(writer, index=False)"""

        self.master.quit()

if __name__ == "__main__":
    import sentry_sdk
    sentry_sdk.init(
        dsn="https://c5822e9079a54ef2b28d4e93a11ebc86@o126149.ingest.sentry.io/4505211315617792",

        # Set traces_sample_rate to 1.0 to capture 100%
        # of transactions for performance monitoring.
        # We recommend adjusting this value in production.
        traces_sample_rate=1.0
    )
    root = tk.Tk()
    app = CompanySelector(root)
    if os.name != "posix":
        sv_ttk.set_theme("dark")
    else:
        sv_ttk.set_theme("dark")
    root.mainloop()
