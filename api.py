from fastapi import FastAPI, File, UploadFile, Form, HTTPException
import tempfile, os, conversor, check_bruto
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

origins = [
    "http://localhost:3000",
    "http://127.0.0.1:3000",
    "http://localhost:5173",
    "http://127.0.0.1:5173",
    "http://localhost:8080",
    "http://127.0.0.1:8080",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,       # se usar cookies/autenticação
    allow_methods=["*"],          # ou liste: ["POST", "GET", "OPTIONS"]
    allow_headers=["*"],          # ou liste os headers usados
)

@app.post("/processar")
async def processar(
    exceis: list[UploadFile] = File(...),
    aba: list[str] = Form(...),
    nome_saida: str = Form("planilhas_agrupadas"),
    modo: str = Form(...),
):
    modo = modo.lower().strip()
    if modo not in {"conferencia", "check"}:
        raise HTTPException(status_code=400, detail="modo inválido. Use 'conferencia' ou 'check'.")

    arquivos_selecionados: dict[str, str] = {}
    temp_paths: list[str] = []

    try:
        for i, excel in enumerate(exceis):
            tab_name = aba[i] if i < len(aba) else f"Unknown-{i+1}"
            contents = await excel.read()
            _, ext = os.path.splitext(excel.filename)
            if not ext:
                ext = ".xlsx"
            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                tmp.write(contents)
                temp_path = tmp.name
            arquivos_selecionados[tab_name] = temp_path
            temp_paths.append(temp_path)
            print(f"{tab_name}: {excel.filename} -> {temp_path}")

        downloads_dir = "/home/massani/Downloads"
        os.makedirs(downloads_dir, exist_ok=True)
        saida_path = os.path.join(downloads_dir, f"{nome_saida}.xlsx")

        # 1) Agrega os arquivos nas abas corretas
        caminho_gerado = conversor.agrupar_excels_em_um(arquivos_selecionados, saida_path)

        caminho_final = check_bruto.api(caminho_gerado, modo)

        return {"message": "OK", "output_path": caminho_final, "modo": modo}

    finally:
        for p in temp_paths:
            try:
                os.unlink(p)
            except Exception:
                pass