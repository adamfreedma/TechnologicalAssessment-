import random
import constants
import re

import json
from openai import OpenAI
import pandas as pd
API_KEY = ""  # Replace with your actual API key
client = OpenAI(api_key=API_KEY)

def merge_consecutive_segments(segments):
    """Merges adjacent segments if they have the same category."""
    if not segments:
        return []
    merged_dict = {}
    for seg in segments:
        cat = seg["category"]
        if cat in merged_dict:
            merged_dict[cat]["text"] += " " + seg["text"]
        else:
            merged_dict[cat] = seg.copy()
    # Return merged segments in the order of their first appearance
    seen = set()
    merged = []
    for seg in segments:
        cat = seg["category"]
        if cat not in seen:
            merged.append(merged_dict[cat])
            seen.add(cat)
    return merged


def get_ai_classif(sentence, is_improve = True):
   
    base_definitions = """
    קטגוריות מותרות: ["מנהיגות", "התנהלות", "מקצועיות", "יכולות תוך אישיות", "יכולות בין אישיות", "אחר"]
    1. מנהיגות: הובלת אנשים, כריזמה וסחיפת אחרים.
    2. התנהלות: תעדוף, לו"ז, משמעת ונהלים.
    3. מקצועיות: ידע, איכות ביצוע, יצירתיות.
    4. יכולות תוך אישיות: ניהול עצמי, ויסות רגשי, בטחון, מוטיבציה.
    5. יכולות בין אישיות: תקשורת, עבודת צוות, רגישות לזולת.
    6. אחר: הערות טכניות או פניות אישיות.
    """

    improve_prompt_content = f"""
    משימה: סווג את המשפט הבא (סוג: {'שיפור/ביקורת' if is_improve else 'שימור/חיובי'}).
    {base_definitions}

    הנחיות קריטיות:
    - חובת סיווג: אסור להחזיר רשימה ריקה. השתמש ב-'אחר' כברירת מחדל.
    - פיצול: פצל לסגמנטים נפרדים רק אם יש שינוי נושא מהותי (למשל מקצועיות מול חברה).
    - איחוד: אל תפצל 'ריכוך' (כמו "אתה טוב אבל...") במשפטי שיפור.

    Text: "{sentence}"
    return json: {{ "segments": [ {{ "text": "חלק המשפט", "category": "שם הקטגוריה" }} ] }}
    """

    response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a helpful assistant that outputs strict JSON."},
                {"role": "user", "content": improve_prompt_content }
            ],
            response_format={"type": "json_object"},
            temperature=0
        )

    data = json.loads(response.choices[0].message.content)
    raw_segments = data.get("segments", [])
    # Apply the merging logic here
    merged_segments = merge_consecutive_segments(raw_segments)
    # Convert Hebrew category to English using the mapping
    for seg in merged_segments:
        heb_cat = seg["category"]
        seg["category"] = HEB_TO_ENG.get(heb_cat, "other")  # Default to "other" if category not found in mapping
    return merged_segments

# Mapping from Hebrew to English category names
HEB_TO_ENG = {
    "מנהיגות": "leadership",
    "התנהלות": "conduct",
    "מקצועיות": "professionalism",
    "יכולות תוך אישיות": "outerpersonal",
    "יכולות בין אישיות": "intrapersonal",
    "אחר": "other"
}
# def extract_real_split(text):
#     # Handles formats like '1. seq1 2. seq2', '1 seq1 2 seq2', etc.
#     match = re.match(r'1\\.?\\s*(.*?)\\s*2\\.?\\s*(.*)', text)
#     if match:
#         return [match.group(1).strip(), match.group(2).strip()]
#     return [text.strip()]
# def main():
#     df = pd.read_excel('splitting.xlsx')
#     # Assume the sentences are in the first column
#     sentences = df.iloc[:, 0].dropna().astype(str)
#     with open('splitting_results.txt', 'w', encoding='utf-8') as f:
#         for idx, sentence in enumerate(sentences, 1):
#             ai_segments = get_ai_classif(sentence)
#             ai_split = [seg['text'] for seg in ai_segments]
#             real_split = extract_real_split(sentence)
#             f.write(f"Sentence {idx}:\n")
#             f.write(f"AI Split: {ai_split}\n")
#             f.write(f"Real Split: {real_split}\n")
#             f.write('-' * 40 + '\n')

# if __name__ == '__main__':
#     main()

# def extract_full_sentence(text):
#     """
#     Extracts the full sentence from a string like '1. seq1 2. seq2' or '1 seq1 2 seq2'.
#     Returns the concatenated sentence as a string.
#     """
#     # Find all sequences after '1' and '2' (with or without dot)
#     matches = re.findall(r'(?:1\\.?\\s*)(.*?)(?:2\\.?\\s*)(.*)', text)
#     if matches:
#         seq1, seq2 = matches[0]
#         return (seq1 + " " + seq2).strip()
#     return text.strip()

