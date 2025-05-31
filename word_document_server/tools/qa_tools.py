"""
Quality Assurance tools for Word Document Server.

These tools extract plain text for external linting and perform simple
style checks like sentence length and rudimentary passive-voice detection.
"""
import os
import re
import json

from docx import Document

from word_document_server.utils.file_utils import ensure_docx_extension


async def extract_text(
    filename: str,
) -> str:
    """Extract all text from a Word document as plain text."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    try:
        doc = Document(filename)
        paragraphs = [p.text for p in doc.paragraphs if p.text]
        text = "\n".join(paragraphs)
        return text
    except Exception as e:
        return f"Failed to extract text: {e}"

async def check_sentence_length(
    filename: str,
    max_chars: int = 120,
) -> str:
    """Find sentences longer than max_chars and report them."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    try:
        text = await extract_text(filename)
        # Split into sentences (simple regex)
        sentences = re.split(r'(?<=[\.\!\?])\s+', text)
        long_sents = [s for s in sentences if len(s) > int(max_chars)]
        return json.dumps(long_sents, indent=2)
    except Exception as e:
        return f"Failed to check sentence length: {e}"

async def check_passive_voice(
    filename: str,
) -> str:
    """Detect simple passive voice patterns and return matching sentences."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    try:
        text = await extract_text(filename)
        # Simple passive voice pattern: forms of 'be' + past participle
        pattern = re.compile(r'\b(?:is|are|was|were|be|been|being)\s+\w+ed\b', re.IGNORECASE)
        sentences = re.split(r'(?<=[\.\!\?])\s+', text)
        passive = [s for s in sentences if pattern.search(s)]
        return json.dumps(passive, indent=2)
    except Exception as e:
        return f"Failed to check passive voice: {e}"