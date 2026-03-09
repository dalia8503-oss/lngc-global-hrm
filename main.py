from fastapi import FastAPI

app = FastAPI()

@app.get("/")
def read_root():
    return {"message": "Hello LNG API"}

#테스트 4 - 터미널에서 내용 수정