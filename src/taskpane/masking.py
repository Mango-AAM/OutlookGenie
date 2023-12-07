import sys
import spacy

def mask_text(text):
    nlp = spacy.load("en_core_web_sm")
    doc = nlp(text)

    # Initialize an empty string for the masked text
    masked_text = ""

    # Keep track of the last index we've processed
    last_idx = 0

    for ent in doc.ents:
        # Add the text before the entity
        masked_text += text[last_idx:ent.start_char]

        # Replace the entity with its label
        masked_text += f"[{ent.label_}]"

        # Update the last index
        last_idx = ent.end_char

    # Add the remaining text after the last entity
    masked_text += text[last_idx:]

    return masked_text

if __name__ == "__main__":
    input_text = sys.argv[1]
    #input_text=input("Enter text: ")
    masked_text = mask_text(input_text)
    print(masked_text)