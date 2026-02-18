from docx.shared import RGBColor

PERSONAL_CATEGORIES = [
    "responsibility",
    "excellence",
    "integrity",
    "daring",
    "mission",
    "fellowship",
    "humility",
]
PROFESSIONAL_CATEGORIES = [
    "knowing",
    "intrapersonal",
    "functioning in society",
    "leadership",
    "conduct",
    "academy",
    "applicative knowledge",
    "security",
]

CATEGORY_NAME_DICT = {
    "knowing": "רמת היכרות",
    "intrapersonal": "יכולות תוך אישיות",
    "functioning in society": "יכולות בין אישיות",
    "leadership": "הובלה",
    "conduct": "ניהול והתנהלות",
    "academy": "מדעי אקדמי",
    "applicative knowledge": "מדעי יישומי",
    "security": "ביטחוני",
    # "commander": "ביקורתיות אפקטיבית",
    "responsibility": "אחריות",
    "excellence": "מצוינות",
    "integrity": "יושרה",
    "daring": "העזה",
    "mission": "שליחות",
    # "fellowship": "רעות",
    # "humility": "ענווה",
    "general": "כללי",
}
OUTPUT_PATH = "output/"

COMMANDERSHIP_CATEGORIES = ["intrapersonal", "functioning in society", "leadership"]
PROFESSIONAL_GRAPH_CATEGORIES = ["academy", "conduct", "applicative knowledge"]

MAIN_CATEGORIES = [
    "intrapersonal",
    "functioning in society",
    "leadership",
    "conduct",
    "academy",
]

WORDED_CATEGORIES = [
    "intrapersonal",
    "outerpersonal",
    "leadership",
    "conduct",
    "professionalism",
    "other"
]

TABLE_CATEGORIES = [
    "knowing",
    "intrapersonal",
    "functioning in society",
    "leadership",
    "conduct",
    "academy",
    "applicative knowledge",
    "security",
    # "commander"
]
KNOWING_CATEGORIES = [
    "intrapersonal",
    "functioning in society",
    "leadership",
    "conduct",
    "academy",
    "applicative knowledge",
    "security",
]
SEMESTER = "ג"


FONTSIZE = 36
SMALL_FONTSIZE = 16

RED_COLOR = RGBColor(255, 0, 0)
GREEN_COLOR = RGBColor(0, 176, 80)
