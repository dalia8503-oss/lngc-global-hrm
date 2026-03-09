from fastapi import FastAPI

app = FastAPI()

@app.get("/")
def read_root():
    return {"message": "Hello LNG API"}

#테스트 4 - 터미널에서 내용 수정
#main.py에서 내용을 수정하고
#ctrl + s로 저장한 후
#터미널에서 git add .
#git commit -m "내용 수정"
#git push
#이렇게 하면 수정한 내용이 원격 저장소에 반영됩니다.