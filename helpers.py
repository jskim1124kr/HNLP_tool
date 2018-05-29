
import re

def isNumber(a):
    try:
        float(a)
        return True
    except ValueError:
        return False


def text_cleaner(text):
    # cleaned_text = re.sub('[\'\',."()]', '', text)
    cleaned_text = re.sub('[.,?/*!]', '', text)
    return cleaned_text