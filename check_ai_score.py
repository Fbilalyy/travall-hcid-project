"""
Extract text from docx and check AI detection score using ZeroGPT free API.
"""
import json
import urllib.request
import urllib.parse
from docx import Document
import sys
import textwrap

def extract_text(docx_path):
    doc = Document(docx_path)
    paragraphs = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            paragraphs.append(text)
    return '\n'.join(paragraphs)

def check_zerogpt(text):
    """Use ZeroGPT free API endpoint"""
    url = "https://api.zerogpt.com/api/detect/detectText"
    payload = json.dumps({"input_text": text}).encode('utf-8')
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0",
        "Origin": "https://www.zerogpt.com",
        "Referer": "https://www.zerogpt.com/"
    }

    req = urllib.request.Request(url, data=payload, headers=headers, method='POST')
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode())
            return data
    except Exception as e:
        return {"error": str(e)}

def analyze_text_features(text):
    """Analyze text features that AI detectors look for"""
    sentences = [s.strip() for s in text.replace('!', '.').replace('?', '.').split('.') if s.strip()]
    words = text.split()

    # Sentence length variance
    sent_lengths = [len(s.split()) for s in sentences]
    avg_len = sum(sent_lengths) / len(sent_lengths) if sent_lengths else 0
    variance = sum((l - avg_len) ** 2 for l in sent_lengths) / len(sent_lengths) if sent_lengths else 0

    # Unique word ratio (lexical diversity)
    unique_ratio = len(set(w.lower() for w in words)) / len(words) if words else 0

    # Average word length
    avg_word_len = sum(len(w) for w in words) / len(words) if words else 0

    # Paragraph count
    paragraphs = [p for p in text.split('\n') if p.strip()]

    # First person usage
    first_person = sum(1 for w in words if w.lower() in ['i', 'my', 'me', 'mine', 'myself', "i'm", "i've", "i'd"])

    # Transition words (informal ones are more human)
    informal_transitions = ['anyway', 'though', 'actually', 'basically', 'honestly',
                          'well,', 'so,', 'but', 'and', 'that said', 'in a sense']
    informal_count = sum(1 for w in informal_transitions if w in text.lower())

    # Contractions (human indicator) - check both straight and curly apostrophes
    normalized = text.lower().replace('\u2019', "'").replace('\u2018', "'").replace('\u201c', '"').replace('\u201d', '"')
    contractions = ["don't", "doesn't", "didn't", "won't", "can't", "isn't", "aren't",
                   "wasn't", "weren't", "it's", "that's", "there's", "they're", "we're",
                   "you're", "i'm", "i've", "i'd", "they'd", "you've", "you'll",
                   "he's", "she's", "we've", "we'd", "couldn't", "wouldn't", "shouldn't",
                   "haven't", "hasn't", "here's"]
    contraction_count = sum(1 for c in contractions if c in normalized)

    print("\n" + "="*60)
    print("METIN ANALIZI - AI TESPIT RISK DEGERLENDIRMESI")
    print("="*60)
    print(f"\nToplam kelime sayisi: {len(words)}")
    print(f"Toplam cumle sayisi: {len(sentences)}")
    print(f"Toplam paragraf sayisi: {len(paragraphs)}")
    print(f"\nOrtalama cumle uzunlugu: {avg_len:.1f} kelime")
    print(f"Cumle uzunlugu varyans: {variance:.1f}")
    print(f"  → {'IYI (degisken)' if variance > 50 else 'RISKLI (cok duzgun)' if variance < 20 else 'ORTA'}")
    print(f"\nLeksikal cesitlilik (unique/total): {unique_ratio:.3f}")
    print(f"  → {'IYI (zengin)' if unique_ratio > 0.5 else 'RISKLI (tekrarci)' if unique_ratio < 0.35 else 'ORTA'}")
    print(f"\nOrtalama kelime uzunlugu: {avg_word_len:.1f} karakter")
    print(f"Birinci sahis kullanimi: {first_person} kez")
    print(f"  → {'IYI (kisisel ton)' if first_person > 10 else 'RISKLI (impersonal)' if first_person < 3 else 'ORTA'}")
    print(f"\nInformal gecis kelimeleri: {informal_count}")
    print(f"  → {'IYI (dogal akis)' if informal_count > 5 else 'RISKLI (formal)' if informal_count < 2 else 'ORTA'}")
    print(f"\nKisaltma (contraction) kullanimi: {contraction_count}")
    print(f"  → {'IYI (konusmaci ton)' if contraction_count > 3 else 'EKLENEBILIR' if contraction_count < 2 else 'ORTA'}")

    # Overall score estimation
    score = 0
    if variance > 50: score += 2
    elif variance > 30: score += 1
    if unique_ratio > 0.5: score += 2
    elif unique_ratio > 0.4: score += 1
    if first_person > 10: score += 2
    elif first_person > 5: score += 1
    if informal_count > 5: score += 2
    elif informal_count > 2: score += 1
    if contraction_count > 3: score += 2
    elif contraction_count > 1: score += 1

    max_score = 10
    human_pct = (score / max_score) * 100

    print(f"\n{'='*60}")
    print(f"TAHMINI INSAN SKORU: {human_pct:.0f}% insan benzeri")
    print(f"{'='*60}")
    if human_pct >= 70:
        print("DURUM: Metin yeterince dogal ve insan benzeri gorunuyor.")
    elif human_pct >= 50:
        print("DURUM: Metin kabul edilebilir ama bazi bolumler AI gibi algilanabilir.")
    else:
        print("DURUM: Metin AI tarafindan yazilmis gibi gorunuyor, duzenleme onerilir.")

    return human_pct

if __name__ == '__main__':
    docx_path = "/Users/fatihbilalyilmaz/PycharmProjects/IE492/IE48L/IE48L_Assignment2.docx"

    print("Metin cikariliyor...")
    text = extract_text(docx_path)
    print(f"Cikarilan metin: {len(text)} karakter, {len(text.split())} kelime\n")

    # Show first 200 chars
    print("Ilk 200 karakter:")
    print(textwrap.shorten(text, width=200, placeholder="..."))

    # Local analysis
    analyze_text_features(text)

    # Try ZeroGPT API
    print(f"\n{'='*60}")
    print("ZEROGPT API KONTROLU")
    print("="*60)

    # Send in chunks if too long (API limit)
    chunk = text[:5000]  # First 5000 chars
    print(f"API'ye gonderiliyor ({len(chunk)} karakter)...")
    result = check_zerogpt(chunk)

    if "error" in result:
        print(f"API hatasi: {result['error']}")
        print("\nAlternatif: gptzero.me veya zerogpt.com sitesine gidip metni yapistirin.")
    else:
        print(f"\nAPI Sonucu:")
        print(json.dumps(result, indent=2, ensure_ascii=False))
        if 'is_human_written' in result:
            print(f"\nInsan tarafindan mi yazildi: {result.get('is_human_written')}")
        if 'ai_percentage' in result:
            print(f"AI yuzdesi: {result.get('ai_percentage')}%")
