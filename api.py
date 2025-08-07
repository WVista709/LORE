from fastapi import FastAPI

app = FastAPI()

@app.get("/")
def read_root():
    return {"message": "Hello World"}

@app.get("/users/{user_id}")
def get_user(user_id: int):
    return {"user_id": user_id, "name": "João"}

@app.post("/users")
def create_user(name: str, email: str):
    return {"name": name, "email": email, "id": 123}