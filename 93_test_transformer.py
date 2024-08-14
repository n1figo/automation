from sentence_transformers import SentenceTransformer, util
   import fitz  # PyMuPDF
   from bs4 import BeautifulSoup
   from playwright.sync_api import sync_playwright

   def get_html_content(url, tab_selector):
       with sync_playwright() as p:
           browser = p.chromium.launch(headless=True)
           page = browser.new_page()
           page.goto(url)
           page.click(tab_selector)
           page.wait_for_load_state('networkidle')
           content = page.content()
           browser.close()
       return content

   def extract_sentences(text):
       return [sent.strip() for sent in text.split('.') if sent.strip()]

   def extract_text_from_html(html_content):
       soup = BeautifulSoup(html_content, 'html.parser')
       return ' '.join(soup.stripped_strings)

   def extract_text_from_pdf(pdf_path):
       doc = fitz.open(pdf_path)
       text = ""
       for page in doc:
           text += page.get_text()
       doc.close()
       return text

   def compare_documents(text1, text2):
       model = SentenceTransformer('paraphrase-MiniLM-L6-v2')
       sentences1 = extract_sentences(text1)
       sentences2 = extract_sentences(text2)
       
       embeddings1 = model.encode(sentences1, convert_to_tensor=True)
       embeddings2 = model.encode(sentences2, convert_to_tensor=True)
       
       cosine_scores = util.pytorch_cos_sim(embeddings1, embeddings2)
       
       differences = []
       for i, sent1 in enumerate(sentences1):
           max_score = max(cosine_scores[i])
           if max_score < 0.8:  # Threshold for considering as different
               differences.append(f"Different: {sent1}")
       
       return differences

   def main():
       url = "https://www.kbinsure.co.kr/CG302120001.ec"
       pdf_path = "path_to_your_pdf_summary.pdf"  # Replace with actual path

       signup_html = get_html_content(url, 'a#tabexmpl')
       coverage_html = get_html_content(url, 'a#tabguarnt')

       signup_text = extract_text_from_html(signup_html)
       coverage_text = extract_text_from_html(coverage_html)
       pdf_text = extract_text_from_pdf(pdf_path)

       print("Differences in Signup tab:")
       print('\n'.join(compare_documents(pdf_text, signup_text)))
       print("\nDifferences in Coverage tab:")
       print('\n'.join(compare_documents(pdf_text, coverage_text)))

   if __name__ == "__main__":
       main()