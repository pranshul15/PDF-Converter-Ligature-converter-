import os
import re
import shutil
from datetime import datetime
from docx import Document
import ollama

def split_large_paragraph(text, max_words=150):
    """Splits a large paragraph into smaller chunks based on sentences."""
    words = text.split()
    if len(words) <= max_words:
        return [text]

    sentences = re.split(r'(?<=[.!?])\s+', text)
    chunks = []
    current_chunk = []
    current_word_count = 0

    for sentence in sentences:
        sentence_words = len(sentence.split())
        if current_word_count + sentence_words > max_words and current_chunk:
            chunks.append(" ".join(current_chunk))
            current_chunk = [sentence]
            current_word_count = sentence_words
        else:
            current_chunk.append(sentence)
            current_word_count += sentence_words

    if current_chunk:
        chunks.append(" ".join(current_chunk))
    return chunks

def split_combined_words_with_ollama(text, model_name="qwen3.5:2b"):
    """Sends text to Ollama strictly to fix accidentally joined words."""
    if not text.strip():
        return text

    prompt = (
        "You are a text repair assistant. Your ONLY job is to identify words that have "
        "been accidentally squashed together without a space and insert the correct space.\n"
        "Example: 'the work exceptionalteam' -> 'the work exceptional team'.\n\n"
        "CRITICAL RULES:\n"
        "1. Do NOT change style, wording, grammar, or punctuation.\n"
        "2. Do NOT fix typos unless it is a missing space between two valid words.\n"
        "3. If there are no merged words, return the original text exactly as it is.\n"
        "4. Output ONLY the repaired text. No explanations, no introduction.\n\n"
        f"Text to process:\n{text}"
    )

    try:
        response = ollama.generate(
            model=model_name,
            prompt=prompt,
            options={"temperature": 0.0}
        )
        return response['response'].strip()
    except Exception as e:
        print(f"Error communicating with Ollama: {e}")
        return text

def merge_spaces_into_runs(paragraph, corrected_text):
    """
    Surgically injects changes (like added spaces) into existing document runs
    to preserve complex inline formatting (bold, italics, etc.).
    """
    runs = paragraph.runs
    if not runs:
        return

    # Map out every single character from the original runs
    orig_chars = []
    for run_idx, run in enumerate(runs):
        for char in run.text:
            orig_chars.append({'char': char, 'run_idx': run_idx})

    # Prepare containers for the new text strings per run
    new_run_texts = ["" for _ in runs]
    orig_ptr = 0

    # Align corrected text back to original runs
    for c in corrected_text:
        if orig_ptr < len(orig_chars) and c == orig_chars[orig_ptr]['char']:
            r_idx = orig_chars[orig_ptr]['run_idx']
            new_run_texts[r_idx] += c
            orig_ptr += 1
        elif c == ' ':
            if orig_ptr < len(orig_chars):
                r_idx = orig_chars[orig_ptr]['run_idx']
            else:
                r_idx = len(runs) - 1
            new_run_texts[r_idx] += c
        else:
            if orig_ptr < len(orig_chars):
                r_idx = orig_chars[orig_ptr]['run_idx']
                new_run_texts[r_idx] += c
                orig_ptr += 1

    # Apply the rebuilt text strings back to the actual runs
    for run_idx, text in enumerate(new_run_texts):
        runs[run_idx].text = text

def create_backup(file_path):
    """Creates a timestamped backup of the file before processing."""
    if not os.path.exists(file_path):
        return None
    
    dir_name, file_name = os.path.split(file_path)
    base_name, ext = os.path.splitext(file_name)
    
    # Generate timestamp: YYYYMMDD_HHMMSS
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file_name = f"{base_name}_backup_{timestamp}{ext}"
    backup_path = os.path.join(dir_name, backup_file_name)
    
    print(f"[Backup] Creating safeguard copy at: {backup_path}")
    shutil.copy2(file_path, backup_path)
    return backup_path

def process_document_in_place(file_path, model_name="qwen3.5:2b", max_words_per_chunk=150):
    """Reads a .docx file, fixes merged words, and overwrites it preserving formatting."""
    if not os.path.exists(file_path):
        print(f"Error: The file '{file_path}' does not exist.")
        return

    # Trigger backup creation first
    backup_path = create_backup(file_path)
    if not backup_path:
        print("Failed to create backup. Aborting operation for safety.")
        return

    print(f"Loading document: {file_path}...")
    doc = Document(file_path)
    total_paragraphs = len(doc.paragraphs)
    print(f"Scanning {total_paragraphs} paragraphs for combined words...")

    changes_made = 0

    for index, paragraph in enumerate(doc.paragraphs, start=1):
        original_text = paragraph.text
        
        if original_text.strip():
            text_chunks = split_large_paragraph(original_text, max_words=max_words_per_chunk)
            corrected_chunks = []
            
            for chunk in text_chunks:
                corrected_chunk = split_combined_words_with_ollama(chunk, model_name)
                corrected_chunks.append(corrected_chunk)
            
            new_text = " ".join(corrected_chunks)
            
            # If changes were detected, merge them back without breaking runs
            if new_text != original_text:
                print(f"  [Fixed] Found and split combined words in paragraph {index}/{total_paragraphs}.")
                merge_spaces_into_runs(paragraph, new_text)
                changes_made += 1

    if changes_made > 0:
        print(f"Saving changes directly back to: {file_path}...")
        doc.save(file_path)
        print(f"Done! Successfully updated {changes_made} paragraphs.")
    else:
        print("No adjustments needed. Document formatting and text are clean.")
        # Optional: You can choose to delete the backup here if no changes happened, 
        # but keeping it is usually safer!

# --- Execution ---
if __name__ == "__main__":
    TARGET_FILE = "sample_document.docx" 
    MODEL = "gemma4:e4b"
    MAX_WORDS = 150 

    process_document_in_place(TARGET_FILE, MODEL, max_words_per_chunk=MAX_WORDS)