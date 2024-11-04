import pptx
import pandas as pd
import re
import os
from bs4 import BeautifulSoup


class_name = "sorcerer/wizard"
spell_list = ["dancing lights", "resistance", "ray of frost", "mage hand", "read magic", "prestidigitation",
              "detect magic", "detect poison", "mage armor", "burning hands", "disguise self", "identify",
              "chastise", "silent image", "fire breath", "flaming sphere", "color spray", "see invisibility"]


def replace_text(text_frame, new_text, p_idx: int = None):
    if p_idx is None:
        p = text_frame.paragraphs[0]
    else:
        p = text_frame.paragraphs[p_idx]

    if p.runs:
        orig_run = p.runs[0]
        # save formatting
        font = orig_run.font
        font_name = font.name
        font_size = font.size
        font_bold = font.bold
        font_italic = font.italic
        font_underline = font.underline
        try:
            font_color = font.color.rgb
            rgbOk = True
        except AttributeError:
            font_color = font.color.theme_color
            rgbOk = False
        # Clear existing text and apply new text with preserved formatting
        if p_idx is None:
            text_frame.clear()  # Clears all text and formatting
            new_run = text_frame.paragraphs[0].add_run()  # New run in first paragraph
        else:
            p.clear()
            new_run = p.add_run()
        new_run.text = new_text
        # Reapply formatting
        new_run.font.name = font_name
        new_run.font.size = font_size
        new_run.font.bold = font_bold
        new_run.font.italic = font_italic
        new_run.font.underline = font_underline
        if rgbOk:
            new_run.font.color.rgb = font_color
        else:
            new_run.font.color.theme_color = font_color
    else:
        text_frame.text = new_text


def replace_text_by_id(slide, textbox_id, new_text, *args):
    shapes = slide.shapes
    if len(shapes) == 0:
        raise ValueError("Empty shapes!")

    for s in shapes:
        if s.shape_id == textbox_id:
            break

    if s.shape_id != textbox_id:
        raise ValueError(f"Text box with id {textbox_id} was not found")

    replace_text(s.text_frame, new_text, *args)


def replace_text_pretty(text_frame, html_format):
    orig_run = text_frame.paragraphs[0].runs[0]
    # save font name and size
    font = orig_run.font
    font_name = font.name
    font_size = font.size

    # clear other formatting
    font_bold = False
    font_italic = False
    font_underline = False
    text_frame.clear()


    try:
        font_color = font.color.rgb
        rgbOk = True
    except AttributeError:
        font_color = font.color.theme_color
        rgbOk = False

    soup = BeautifulSoup(html_format, 'html.parser')
    current_node = next(next(soup.children).children) # we know that soup contains at least 1 paragraph, which is already added
    visit_stack = []
    for child in reversed(soup.contents[1:]):
        visit_stack.append((child, True))
    for child in reversed(soup.contents[0].contents):
        visit_stack.append((child, True))
    while visit_stack:
        current_node, is_before = visit_stack.pop()
        if current_node.name is None:    # current node is a leaf text
            new_run = text_frame.paragraphs[-1].add_run()
            new_run.text = str(current_node)
            new_run.font.name = font_name
            new_run.font.size = font_size
            new_run.font.bold = font_bold
            new_run.font.italic = font_italic
            new_run.font.underline = font_underline
        elif current_node.name == 'p':
            text_frame.add_paragraph()
            for child in reversed(current_node.contents):
                visit_stack.append((child, True))
        else:
            if current_node.name == 'i':
                font_italic = is_before
            elif current_node.name == 'b':
                font_bold = is_before
            elif current_node.name == 'u':
                font_underline = is_before
            else:
                continue

            if is_before:
                visit_stack.append((current_node, False))
                for child in reversed(current_node.contents):
                    visit_stack.append((child, True))


if __name__ == "__main__":
    print("Loading spell database")
    df = pd.read_excel("spells_by_class.xlsx", "Sorcerer-Wizard")
    print("Loading done")

    spell_list = pd.Series(spell_list)
    spell_list = spell_list.str.lower()

    lower_names = df['name'].str.lower()
    selected_spells = df.loc[lower_names.isin(spell_list)]
    found_num = selected_spells.shape[0]
    print(f"{found_num}/{spell_list.size} spells have been found.")
    print("Initiate export")
    if not os.path.isdir('./output'):
        print("Output folder not found, create it")
        os.mkdir('./output')

    for cnt, idx_row in enumerate(selected_spells.iterrows()):
        spell_row = idx_row[1]
        presentation = pptx.Presentation(f"templates/{spell_row['school']}_template.pptx")
        shapes = presentation.slides[0].shapes  # Shape order: background pic, name, components, table, level

        replace_text_by_id(presentation.slides[0], 7, spell_row['name'], 0)  # id: 7
        replace_text_by_id(presentation.slides[0], 7, spell_row['components'], 1)      # id: 8

        table = shapes[2].table
        replace_text(table.cell(0, 1).text_frame, spell_row['casting_time'])
        replace_text(table.cell(0, 3).text_frame, spell_row['duration'])
        replace_text(table.cell(1, 1).text_frame, spell_row['range'])
        if (not pd.isna(spell_row['targets'])) and (not pd.isna(spell_row['area'])):
            if spell_row['targets'] == spell_row['area']:
                replace_text(table.cell(1, 2).text_frame, "Area/Target")
                replace_text(table.cell(1, 3).text_frame, spell_row['targets'])
            else:
                raise Exception("What to do if both targets and area is defined?")
        if not pd.isna(spell_row['targets']):
            replace_text(table.cell(1, 2).text_frame, "Target")
            replace_text(table.cell(1, 3).text_frame, spell_row['targets'])
        elif not pd.isna(spell_row['area']):
            replace_text(table.cell(1, 2).text_frame, "Area")
            replace_text(table.cell(1, 3).text_frame, spell_row['area'])
        else:
            table.cell(1, 2).text_frame.clear()
            table.cell(1, 3).text_frame.clear()
        # replace_text(table.cell(2, 0).text_frame, spell_row['description'])
        replace_text_pretty(table.cell(2, 0).text_frame, spell_row['description_formated'])
        lvl = re.search(f"{class_name} [0-9]+", spell_row['spell_level']).group()[-1].strip()
        replace_text_by_id(presentation.slides[0], 11, lvl)     # id: 11

        presentation.save(f"output/{spell_row['name']}.pptx")
        print(f"{cnt+1}/{found_num} done")

    print("Individual export done, create concatenated")


    print("Exporting done")
    not_found_spells = spell_list[~spell_list.isin(lower_names)]
    if not not_found_spells.empty:
        print("The following spells have not been found:")
        print(not_found_spells)
