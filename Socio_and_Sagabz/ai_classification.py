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
    improve_keywords_str = """
    ### defenitions: 
 מנהיגות - היכולת להנחות, לרתום ולהניע להשראה ולהוביל קבוצה או ארגון לקראת השגת מטרה. יכולת ליזום ולהתסכל על הדברים בצורה מערכתית, ולנהל תהליכים.
ניהול והתנהלות - התנהלות נכונה מתייחסת למנגנונים פרקטיים לקבלת החלטות וניהול משאבים נכון. כמו כן לסדר וארגון, יהול זמנים וחלוקת משימות לפי תעדוף נכון. מנסה להרחיב תהליכים בהם הוא מעורב ולשפר ולהשתפר בתוך תפקידו.
מקצועיות - היכולת לבצע משימות ברמה גבוהה, שימוש בידע ומיומנויות, הרחבה והעמקה של הידע שאותו אתה לומד. סקרנות בלמידה ובהבנה, יכולת חשיבה יצירתית.  
יכולות תוך אישיות - היכולות הפנימיות של האדם מסתכל על ערכים תכונות, יכולת לנהל את עצמך, לזהות ולווסת את רגשותיך, ולהשתמש בערך עצמי והבנה לצמיחה ושיפור עצמי מתמיד.
יכולות בין אישיות - מתייחסות לכישורים ומיומנויות המאפשרים לאדם לתקשר, לשתף פעולה ולנהל יחסים טובים עם אחרים וסביבתו. יכולת לנהל דיון קונפליקטים ומורכבויות בתוך קבוצה.
    ### Keyword Dictionary (Strong Indicators):
    Use these keywords to disambiguate:
    1. **התנהלות (Conduct/Execution):** Keywords: "ניהול", "לנהל", "תכנון", "לו\"ז", "יומן", "סדר", "ארגון", "תעדוף", "יעיל", "חפשן", "מפוזר", "מבולגן".
    2. **מקצועיות (Professionalism):** Keywords: "מקצועי", "ידע", "חריצות", "השקעה", "לומד", "ציונים", "אקדמיה",  "מוסר עבודה".
    3. **יכולות בין אישיות (Interpersonal):** Keywords: "חבר", "נעים", ,"שחצן","מתנשא", "צוות",  "אחרים","עזרה", "להיפתח", "אכפתי", "מערכת יחסים".
    4. **יכולות תוך אישיות (Intrapersonal):** Keywords: "שקט", "ביישן", "ביטחון עצמי", "הומור", ,"ציני", "מודעות עצמית", "לקחת ללב", "ענווה".
    5. **מנהיגות (Leadership):** Keywords: "מוביל", "מנהיג", "כריזמה", "פיקוד", "יוזמה". (Focus on leading people or taking initiative).
    """

    conserve_keywords_str = """
    ### defenitions: 
מנהיגות - היכולת להנחות, לרתום ולהניע להשראה ולהוביל קבוצה או ארגון לקראת השגת מטרה. יכולת ליזום ולהתסכל על הדברים בצורה מערכתית, ולנהל תהליכים.
ניהול והתנהלות - התנהלות נכונה מתייחסת למנגנונים פרקטיים לקבלת החלטות וניהול משאבים נכון. כמו כן לסדר וארגון, יהול זמנים וחלוקת משימות לפי תעדוף נכון. מנסה להרחיב תהליכים בהם הוא מעורב ולשפר ולהשתפר בתוך תפקידו.
מקצועיות - היכולת לבצע משימות ברמה גבוהה, שימוש בידע ומיומנויות, הרחבה והעמקה של הידע שאותו אתה לומד. סקרנות בלמידה ובהבנה, יכולת חשיבה יצירתית.  
יכולות תוך אישיות - היכולות הפנימיות של האדם מסתכל על ערכים תכונות, יכולת לנהל את עצמך, לזהות ולווסת את רגשותיך, ולהשתמש בערך עצמי והבנה לצמיחה ושיפור עצמי מתמיד.
יכולות בין אישיות - מתייחסות לכישורים ומיומנויות המאפשרים לאדם לתקשר, לשתף פעולה ולנהל יחסים טובים עם אחרים וסביבתו. יכולת לנהל דיון קונפליקטים ומורכבויות בתוך קבוצה.

    ### Keyword Dictionary (Strong Indicators):
    Use these keywords to disambiguate:
    1. **התנהלות (Conduct/Execution):** Keywords: "ניהול", "לנהל", "תכנון", "לו\"ז", "יומן", "סדר", "ארגון", "תעדוף", "יעיל", "ביצועיסט", "מתקתק", "כנס".
    2. **מקצועיות (Professionalism):** Keywords: "מקצועי", "ידע", "רמה גבוהה", "חריצות", "השקעה", "לומד", "ציונים", "אקדמיה",  (Negative), "מוסר עבודה".
    3. **יכולות בין אישיות (Interpersonal):** Keywords: "חבר", "נעים", "רגיש", "צוות", "עזרה", "קשוב", "אכפתי", "מערכת יחסים".
    4. **יכולות תוך אישיות (Intrapersonal):** Keywords: "שקט", "ביישן", "ביטחון עצמי", "הומור", "ציני", "מודעות עצמית", "לקחת ללב", "ענווה".
    5. **מנהיגות (Leadership):** Keywords: "מוביל", "מנהיג", "סוחף", "כריזמה", "פיקוד", "יוזמה". (Focus on leading people or taking initiative).
    """


    system_instruction = """
    You are an expert sociometric analyst. 
    Your goal is to classify feedback sentences into 5 specific categories.
    **Guideline:** Prefer holistic classification. Do not split the sentence unless there is a clear, distinct contradiction in topics.
    """

    user_prompt_conserve = f"""
    Analyze the following conserve feedback: "{sentence}"

    {conserve_keywords_str}

    ### Logic Flow:
    1. **Check Keywords and Defenitions:** Does the sentence contain specific keywords from the list above or match the definitions?
    2. **Holistic View:** Can the whole sentence fit into one main category? If yes, do not split.
    3. **Split Condition:** Split only if one part clearly discusses Skill X and another distinct part discusses Skill Y.

    The allowed categories are {HEB_TO_ENG.keys()} (in Hebrew). If the sentence does not fit any category, classify it as "אחר" (other).

    ### Output Format:
    Return valid JSON only: {{ "segments": [ {{ "text": "...", "category": "..." }} ] }}
    """


    user_prompt_improve = f"""
    Analyze the following improvement feedback: "{sentence}"

    {improve_keywords_str}

    ### Logic Flow:
    1. **Check Keywords and Defenitions:** Does the sentence contain specific keywords from the list above or match the definitions?
    2. **Holistic View:** Can the whole sentence fit into one main category? If yes, do not split.
    3. **Split Condition:** Split only if one part clearly discusses Skill X and another distinct part discusses Skill Y.

    The allowed categories are {HEB_TO_ENG.keys()} (in Hebrew). If the sentence does not fit any category, classify it as "אחר" (other).

    ### Output Format:
    Return valid JSON only: {{ "segments": [ {{ "text": "...", "category": "..." }} ] }}
    """
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": system_instruction},
            {"role": "user", "content": user_prompt_improve if is_improve else user_prompt_conserve}
        ],
        response_format={"type": "json_object"},
        temperature=0
    )

    content = response.choices[0].message.content
    if content is None:
        # Error handling: return default segment
        return [{"text": sentence, "category": "other"}]
    data = json.loads(content)
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
