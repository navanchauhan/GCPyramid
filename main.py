import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from datetime import datetime

from ttkwidgets import ScrolledListbox

import sv_ttk
import darkdetect

import os
from os import path

from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

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

    for _ in range(4):
        if len(pyramid) > 2:
            if len(pyramid[-1]) < len(pyramid[-2]):
                pyramid[-3].append(pyramid[-2][-1])
                pyramid[-2].pop(-1)

    for _ in range(3):
        if len(pyramid) > 4:
            if len(pyramid[-2]) < len(pyramid[-3]):
                pyramid[-4].append(pyramid[-3][0])
                pyramid[-3].pop(0)

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
        self.master.resizable(width=False, height=False)

        width=600
        height=500
        screenwidth = self.master.winfo_screenwidth()
        screenheight = self.master.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.master.geometry(alignstr)
        self.master.title("UCS Pyramid Builder")


        self.company_listbox = ScrolledListbox(self.master, selectmode=tk.MULTIPLE, exportselection=False, height=20)
        self.company_listbox.place(x=10,y=20,width=235,height=457)
        #self.company_listbox.pack(expand=True)

        for ticker, company in zip(self.df["Company Name"],self.df["Symbol"]):
            self.company_listbox.listbox.insert(tk.END, f' {company} - {ticker}')


        # bind the listbox to an onselect event
        self.company_listbox.listbox.bind('<<ListboxSelect>>', self.onselect)

        self.num_companies_selected_label = ttk.Label(self.master, text="0", font='SunValleyBodyStrongFont 18 bold')
        self.num_companies_selected_label.place(x=315,y=320,width=30,height=30)

        self.companies_label = ttk.Label(self.master, text="Companies Selected", font='SunValleyBodyStrongFont 18')
        self.companies_label.place(x=340,y=320,width=200, height=30)

        self.pyramid_title_string = tk.StringVar()
        self.pyramid_title_string.set(f"Copyright (2004-{datetime.today().year}) by Gentry Capital Corporation")

        #self.pyramid_title =  ttk.Entry(self.master, textvariable=self.pyramid_title_string)
        #self.pyramid_title.pack()

        self.customization_label = ttk.Label(self.master, text="Customization", font='SunValleyBodyStrongFont 12 bold')
        self.customization_label.place(x=250, y=20, width=100, height=20)

        self.options_label = ttk.Label(self.master, text="Options", font='SunValleyBodyStrongFont 12 bold')
        self.options_label.place(x=250, y=150, width=100, height=20)

        self.prepared_for_string = tk.StringVar()
        self.prepared_for_entry = ttk.Entry(self.master, textvariable=self.prepared_for_string)
        self.prepared_for_entry.place(x=380,y=50,width=195,height=30)

        self.prepared_for_label = ttk.Label(self.master, text="Prepared For")
        self.prepared_for_label.place(x=250,y=50,width=132,height=30)

        self.advisor_string = tk.StringVar()
        self.advisor_entry = ttk.Entry(self.master, textvariable=self.advisor_string)
        self.advisor_entry.place(x=380,y=90,width=195,height=30)
        self.advisor_label = ttk.Label(self.master, text="Advisor")
        self.advisor_label.place(x=250,y=90,width=132,height=30)

        self.generate_pyramid_var = tk.IntVar()
        self.generate_pyramid_var.set(1)
        self.generate_pyramid_checkbx = ttk.Checkbutton(self.master, text="Generate Pyramid Image", onvalue=1, offvalue=0, variable=self.generate_pyramid_var)
        self.generate_pyramid_checkbx.place(x=250,y=180, width=329, height=30)

        self.generate_word_doc = tk.IntVar()
        self.generate_word_doc.set(0)
        self.generate_word_doc_checkbx = ttk.Checkbutton(self.master, text="Generate Description Document", onvalue=1, offvalue=0, variable=self.generate_word_doc)
        self.generate_word_doc_checkbx.place(x=250,y=220, width=329, height=30)

        self.create_pyramid_button = ttk.Button(self.master, text="Create", command=self.create_pyramid)
        #self.create_pyramid_button.pack()
        self.create_pyramid_button.place(x=250,y=450,width=150,height=30)

        self.reset_button = ttk.Button(self.master, text="Reset", command=self.reset_fields)
        self.reset_button.place(x=250+150+25, y=450, width=150, height=30)

    def onselect(self, evt):
        w = evt.widget
        num_selected = len(w.curselection())
        self.num_companies_selected_label.config(text=f"{num_selected}")

    def reset_fields(self):
        self.generate_word_doc.set(0)
        self.generate_pyramid_var.set(1)
        self.prepared_for_string.set("")
        self.advisor_string.set("")
        self.company_listbox.listbox.selection_clear(0,'end')
        self.num_companies_selected_label.config(text="0")

    def create_pyramid(self):
        selected_companies_tmp = [self.company_listbox.listbox.get(index) for index in self.company_listbox.listbox.curselection()]
        print(f'Selected {len(selected_companies_tmp)} companies')
        selected_companies = [x.split(" - ")[0].strip() for x in selected_companies_tmp]
        print(f'Symbols: {len(selected_companies)}')
        selected_df = self.df[self.df["Symbol"].isin(selected_companies)].sort_values(by="Weighting", ascending=False)
        print(f'DF: {len(selected_df)}')
        #pyramid_file_path = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")])

        if self.generate_word_doc.get() == 1:
            word_file_path = asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])

            if not word_file_path:
                return

            document = Document("assets/doc_template.docx")

            advisor = self.advisor_string.get()
            today = datetime.today()

            document.paragraphs[0].text = document.paragraphs[0].text.replace("{{ADVISOR}}", advisor)

            to_replace = document.paragraphs[3].text

            to_replace = to_replace.replace("{{MONTH}}" , today.strftime("%B"))
            to_replace = to_replace.replace("{{YEAR}}" , today.strftime("%Y"))

            document.paragraphs[3].text = to_replace

            for company, ticker, dividend, currency, desc in zip(selected_df["Company Name"], selected_df["Symbol"], selected_df["Dividend?"], selected_df["Currency"], selected_df["Blurb"]):
                extra_char = ""
                if dividend.strip() == "Y":
                    extra_char = "***"

                new_paragraph = document.add_paragraph()
                temp_run = new_paragraph.add_run(f'{company}{extra_char} ({ticker})')
                temp_run.bold = True
                new_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                document.paragraphs.pop()
                desc_paragraph = document.add_paragraph(f"\n{desc}\n")
                desc_paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                desc_paragraph.bold = False

            document.save(word_file_path)

        if self.generate_pyramid_var.get() == 0:
            None
        else: 
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
            pyramid_width = 3000
            img_height = 1550 #2550
            overall_height = 2550
            block_color = (3, 37, 126)  # Blue color
            dividend_color = block_color #(1,50,32)  # Green color
            font_color = (255, 255, 255)  # White color
            normal_block = (3, 37, 126) # Dark Blue  
            padding = 15  # Padding around blocks
            radius = 10 
            logo_padding = 100 # top padding for logo
            pyramid_padding_bottom = 350 # bottom padding for pyramid
            max_companies_in_row = 10
            base_path = "assets"
            primary_font = path.join(base_path, "baskerville.ttf")
            secondary_font = path.join(base_path, "gill_sans_bold.ttf")

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
            print(f"Still have {len(companies)}")

            companies = pyramid_list(companies)

            for company_row in companies:
                print(company_row)

            # Calculate the maximum number of companies in a group (this will be the width of your pyramid)
            max_companies = max(len(row) for row in companies) #max(len(v) for v in pyramid.values())

            # Calculate the total number of groups (the height of your pyramid)
            total_groups = len(companies) #len(pyramid)

            # Calculate the size of each block based on the width and height of the image and the number of blocks
            block_size = min((pyramid_width - padding) // max_companies - padding, (img_height - padding) // total_groups - padding)

            # Calculate the size of each block based on the width and height of the image and the number of blocks
            block_width = (pyramid_width - padding) // max_companies - padding
            block_height = (img_height - padding) // total_groups - padding

            # Calculate the total width and height of the blocks (including padding)
            total_width = max_companies * (block_width + padding)
            total_height = total_groups * (block_height + padding)

            # Calculate the starting position for the first block
            start_x = (pyramid_width - total_width) // 2
            start_y = (img_height - total_height) // 2


            # Create an image big enough to hold the pyramid
            img = Image.new('RGB', (pyramid_width, img_height), "white")
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
            text4 = datetime.today().strftime("%B %d, %Y")
            text5 = ""

            if self.prepared_for_string.get() != "":
                print(f"Prepared for string present - {self.prepared_for_string.get()}")
                text5 += f"Prepared for {self.prepared_for_string.get()}"
            if self.advisor_string.get() != "":
                advisor = self.advisor_string.get()
                if text5 != "":
                    tmp_var = [text5, f"By {advisor}"]
                    text5 = tmp_var
                else:
                    text5 = f"By {advisor}"

            print(text5)
            draw = ImageDraw.Draw(final_img)
            font_size = 60

            font1 = ImageFont.truetype(primary_font, font_size)
            text_width1 = draw.textlength(text1, font=font1)
            while text_width1 > final_img.width:
                font_size -= 1
                font1 = ImageFont.truetype(primary_font, font_size)
                text_width1 = draw.textlength(text1, font=font1)

            font2_size = font_size - 10
            font2 = ImageFont.truetype(primary_font, font2_size)
            text_width2 = draw.textlength(text2, font=font2)
            while text_width2 > final_img.width:
                font_size -= 1
                font2 = ImageFont.truetype(primary_font, font2_size)
                text_width2 = draw.textlength(text2, font=font2)

            new_font_size = 30
            font3 = ImageFont.truetype(secondary_font, new_font_size)
            _,t,_,b = draw.textbbox((100,100), text3, font=font3)
            while (b-t) < (0.4*pyramid_padding_bottom):
                new_font_size += 1
                font3 = ImageFont.truetype(secondary_font, new_font_size)
                _,t,_,b = draw.textbbox((100,100), text3, font=font3)
            text_width3 = draw.textlength(text3, font=font3)
            while text_width3 > (final_img.width*0.8):
                new_font_size -= 1
                font3 = ImageFont.truetype(secondary_font, new_font_size)
                text_width3 = draw.textlength(text3, font=font3)
            _,t,_,b = draw.textbbox((100,100),text3,font=font3)

            font4_size = font2_size - 5
            font4 = ImageFont.truetype(primary_font, font4_size)
            text_width4 = draw.textlength(text4, font=font4)


            start_x1 = (final_img.width - text_width1) // 2
            start_y1 = image_logo.height + logo_padding + 50

            start_x2 = (final_img.width - text_width2) // 2
            start_y2 = image_logo.height + logo_padding + 125

            start_x3 = (final_img.width - text_width3) // 2
            start_y3 = (final_img.height - pyramid_padding_bottom + ((pyramid_padding_bottom)-(b-t))//2)

            start_x4 = (final_img.width - text_width4) - 150
            start_y4 = logo_padding + 30 

            start_x5 = 150
            start_y5 = start_y4
            
            draw.text((start_x1, start_y1), text1, font=font1, fill="black")
            draw.text((start_x2, start_y2), text2, font=font2, fill="black")
            draw.text((start_x3, start_y3), text3, font=font3, fill=(158,161,162))
            draw.text((start_x4, start_y4), text4, font=font4, fill="black")
            if type(text5) == str:
                draw.text((start_x5, start_y5), text5, font=font4, fill="black")
            else:
                for idx, text2write in enumerate(text5):
                    draw.text((start_x5, start_y5 + (idx*69)), text2write, font=font4, fill="black")

            final_img.paste(img, ((img_width-pyramid_width)//2, l_height + (overall_height - l_height - img_height) - pyramid_padding_bottom) )

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
            final_img.show()
            #img.show()

            """if not pyramid_file_path:
                                        return
                            
                                    with pd.ExcelWriter(pyramid_file_path) as writer:
                                        selected_df.to_excel(writer, index=False)"""

            #self.master.quit()

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
    sv_ttk.set_theme("dark")
    if darkdetect.isLight():
        sv_ttk.set_theme("light")

    root.mainloop()
