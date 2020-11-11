import json, time
from docx import Document, shared

RED = shared.RGBColor(0x7D, 0x08, 0x00)
ORANGE = shared.RGBColor(0x66, 0x58, 0x00)
BLACK = shared.RGBColor(0x00, 0x00, 0x00)


def load(filepath):
    with open(filepath) as infile:
        return json.load(infile)


def get_words_by_start_time(transcript):
    """Merges punctuation with standard words since they don't have a start time, 
    returns them in a handy map of start_time to word and confidence

    Args:
        transcript: Amazon Transcript JSON

    Returns:
        (dict): a map of start_time to word and confidence
    """
    merged_words = {}

    items = transcript["results"]["items"]

    for i, item in enumerate(items):
        # Only save pronunciations... may not be necessary, or may need other types
        if item["type"] == "pronunciation":
            word = item["alternatives"][0]["content"]
            confidence = item["alternatives"][0]["confidence"]

            # If the next item in the transcript is a punctuation, merge with the current word
            if i < len(items) - 1 and items[i + 1]["type"] == "punctuation":
                word += items[i + 1]["alternatives"][0]["content"]

            # Add the word to the map at start time
            merged_words[item["start_time"]] = {
                "content": word,
                "confidence": confidence,
            }

    return merged_words


def build_docx(title, word_time_map, speaker_segments, speaker_names):
    """Builds a Word document version of the transcript, with words color-coded
    according to confidence 

    Args:
        word_time_map (dict): a map of start_time to word and confidence
        transcript (dict): Amazon Transcribe transcript
        speaker_names (dict): a mapping of speaker_label (e.g. spk_0) to name 
            which will appear in transcript
        
    Returns:
        Document: a complete Word document
    """
    doc = Document()
    doc.add_heading(title, 0)

    current_speaker = ""
    current_paragraph = None

    # Start running through each speaker's segment
    for speaker_segment in transcript["results"]["speaker_labels"]["segments"]:
        name = speaker_names[speaker_segment["speaker_label"]]

        # Speaker has changed, add a new sub-heading for with their name and time
        if current_speaker != name:
            speaker_start = speaker_segment["start_time"]
            # Humanise time
            start = time.strftime("%Hh%Mm%Ss", time.gmtime(float(speaker_start)))

            # Begin a new paragraph for the speaker
            doc.add_heading(f"{name} @ {start}", 2)
            current_speaker = name
            current_paragraph = doc.add_paragraph()

            # Capitalise first word for new speaker
            word_time_map[speaker_start]["content"] = word_time_map[speaker_start][
                "content"
            ].capitalize()

        # Build paragraph
        for item in speaker_segment["items"]:
            start = item["start_time"]
            word = word_time_map[start]["content"] + " "

            # Add word, coloured by confidence level
            confidence = float(word_time_map[start]["confidence"])
            current_paragraph.add_run(word).font.color.rgb = (
                RED if confidence < 0.5 else ORANGE if confidence < 0.85 else BLACK
            )

    return doc


if __name__ == "__main__":
    transcript = load("./some_path.json")
    words_by_time = get_words_by_start_time(transcript)

    title = "Document Title"
    name_map = {"spk_0": "Speaker One", "spk_1": "Speaker Two"}

    document = build_docx(title, words_by_time, transcript, name_map)
    document.save(f"{title}.docx")
