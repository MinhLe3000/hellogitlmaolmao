from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import JSONResponse
import pandas as pd
from io import BytesIO
import torch
from transformers import AutoModel, AutoTokenizer
from sklearn.metrics.pairwise import cosine_similarity
try:
    from underthesea import word_tokenize
    USE_UNDERTHESEA = True
except ImportError:
    USE_UNDERTHESEA = False
    print("Thư viện underthesea không được tìm thấy. Sử dụng phương pháp tokenize đơn giản.")

app = FastAPI()
# code test git 
class DepartmentClassifier:
    def __init__(self, keywords_file):
        self.tokenizer = AutoTokenizer.from_pretrained("vinai/phobert-base", use_auth_token=False)
        self.model = AutoModel.from_pretrained("vinai/phobert-base", use_auth_token=False)
        self.departments = self._load_keywords(keywords_file)
        self.department_embeddings = {}
        for dept, keywords in self.departments.items():
            dept_emb = self._get_embeddings(" ".join(keywords))
            self.department_embeddings[dept] = dept_emb.mean(axis=0)

    def _load_keywords(self, excel_file):
        try:
            df = pd.read_excel(excel_file)
            departments = {}
            for column in df.columns:
                keywords = df[column].dropna().tolist()
                keywords = keywords[1:]
                departments[column.lower()] = keywords
            return departments
        except Exception as e:
            print(f"Lỗi khi đọc file keywords: {str(e)}")
            return {}

    def _tokenize_text(self, text):
        if USE_UNDERTHESEA:
            words = word_tokenize(text)
            return " ".join(words)
        else:
            words = text.lower().split()
            return " ".join(words)

    def _get_embeddings(self, text):
        tokenized_text = self._tokenize_text(text)
        encoded = self.tokenizer(tokenized_text, return_tensors='pt', padding=True, truncation=True)
        with torch.no_grad():
            outputs = self.model(**encoded)
            embeddings = outputs.last_hidden_state.numpy()
        return embeddings[0]

    def classify_text(self, text, threshold=0.3):
        text_emb = self._get_embeddings(text)
        text_emb_mean = text_emb.mean(axis=0)
        results = {}
        for dept, dept_emb in self.department_embeddings.items():
            similarity = cosine_similarity([text_emb_mean], [dept_emb])[0][0]
            if similarity >= threshold:
                results[dept] = float(similarity)
        return dict(sorted(results.items(), key=lambda x: x[1], reverse=True))

    def get_similarity_categories(self, classifications):
        very_high = []
        high = []

        for dept, score in classifications.items():
            similarity_percentage = score * 100
            if similarity_percentage > 70:
                very_high.append(dept.upper())
            elif similarity_percentage > 60:
                high.append(dept.upper())

        return very_high, high

@app.post("/classify")
async def classify_file(texts_file: UploadFile = File(...), keywords_file: UploadFile = File(...)):
    try:
        texts_content = await texts_file.read()
        keywords_content = await keywords_file.read()

        texts_df = pd.read_excel(BytesIO(texts_content))
        keywords_path = BytesIO(keywords_content)

        if 'text' not in texts_df.columns:
            raise HTTPException(status_code=400, detail="Không tìm thấy cột 'text' trong file upload")

        classifier = DepartmentClassifier(keywords_path)

        results = []
        for idx, row in texts_df.iterrows():
            text = row['text']
            classifications = classifier.classify_text(text)
            very_high, high = classifier.get_similarity_categories(classifications)

            if very_high:
                department_label = f"Rất cao: {', '.join(very_high)}"
            elif high:
                department_label = f"Cao: {', '.join(high)}"
            else:
                department_label = "Đoạn văn không thuộc phòng ban nào"

            result_row = {
                'text': text
            }
            if 'sentiment' in texts_df.columns:
                result_row['sentiment'] = row['sentiment']

            result_row['departments'] = department_label
            results.append(result_row)

        # Chuyển kết quả sang DataFrame
        results_df = pd.DataFrame(results)

        # Lưu kết quả ra file Excel
        output_file = "classified_results.xlsx"
        results_df.to_excel(output_file, index=False)

        return JSONResponse(content={
            "message": "Classification completed successfully.",
            "output_file": output_file,
            "data": results
        })

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi khi xử lý file: {str(e)}")
