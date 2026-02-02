"""
FastAPI REST API para Processamento de Excel
Expõe endpoints para manipulação de arquivos Excel através de HTTP.
"""
import os
import logging
from fastapi import FastAPI, UploadFile, File, HTTPException, Header, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from services.excel_processor import process_excel

# Configuração de segurança
API_KEY = os.getenv('API_KEY')

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Inicialização do FastAPI
app = FastAPI(
    title="Excel Processing API",
    description="API REST para processamento automatizado de arquivos Excel com regras de negócio financeiras",
    version="1.0.0"
)

# Configuração CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Para desenvolvimento - em produção, especificar domínios permitidos
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# =====================
# Dependências de Segurança
# =====================
async def verify_api_key(x_api_key: str = Header(None)):
    """
    Verifica se a API Key fornecida no header x-api-key é válida.
    
    Args:
        x_api_key: Valor do header x-api-key da requisição
        
    Raises:
        HTTPException 401: Se a chave não for fornecida ou for inválida
    """
    if not API_KEY:
        logger.warning("API_KEY não configurada no ambiente. Endpoint desprotegido!")
        return  # Se não há API_KEY configurada, permite acesso (desenvolvimento)
    
    if not x_api_key:
        logger.warning("Tentativa de acesso sem API Key")
        raise HTTPException(
            status_code=401,
            detail="API Key não fornecida. Inclua o header 'x-api-key' na requisição."
        )
    
    if x_api_key != API_KEY:
        logger.warning(f"Tentativa de acesso com API Key inválida: {x_api_key[:8]}...")
        raise HTTPException(
            status_code=401,
            detail="API Key inválida"
        )
    
    logger.debug("API Key validada com sucesso")
    return x_api_key


@app.get("/")
async def root():
    """
    Endpoint raiz com informações sobre a API.
    """
    return {
        "message": "Excel Processing API",
        "version": "1.0.0",
        "endpoints": {
            "health": "GET /health",
            "process": "POST /process",
            "docs": "GET /docs"
        }
    }


@app.get("/health")
async def health_check():
    """
    Health check endpoint para monitoramento.
    """
    return {
        "status": "healthy",
        "service": "excel-processing-api"
    }


@app.post("/process")
async def process_file(
    file: UploadFile = File(...),
    api_key: str = Depends(verify_api_key)
):
    """
    Processa um arquivo Excel aplicando regras de negócio.
    
    Args:
        file: Arquivo Excel (.xlsx) para processamento
        api_key: API Key validada (via dependency injection)
        
    Returns:
        StreamingResponse com o arquivo processado para download
        
    Raises:
        HTTPException 401: Se a API Key não for fornecida ou for inválida
        HTTPException 400: Se o formato do arquivo não for .xlsx
        HTTPException 500: Se ocorrer erro durante o processamento
    """
    # Validação de tipo de arquivo
    if not file.filename.endswith('.xlsx'):
        logger.warning(f"Tentativa de upload com arquivo inválido: {file.filename}")
        raise HTTPException(
            status_code=400,
            detail="Apenas arquivos .xlsx são suportados"
        )
    
    logger.info(f"Iniciando processamento do arquivo: {file.filename}")
    
    try:
        # Leitura do arquivo
        file_bytes = await file.read()
        logger.info(f"Arquivo lido com sucesso: {len(file_bytes)} bytes")
        
        # Processamento
        output = process_excel(file_bytes)
        logger.info(f"Processamento concluído com sucesso para: {file.filename}")
        
        # Preparação da resposta
        output_filename = f"processado_{file.filename}"
        
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f'attachment; filename="{output_filename}"'
            }
        )
        
    except ValueError as e:
        # Erros de validação de dados/estrutura do Excel
        logger.error(f"Erro de validação ao processar {file.filename}: {e}")
        print(f"ERRO DE VALIDAÇÃO: {e}")  # Print adicional para debug no terminal
        raise HTTPException(
            status_code=400,
            detail=f"Erro de validação: {str(e)}"
        )
        
    except Exception as e:
        # Outros erros de processamento
        logger.error(f"Erro inesperado ao processar {file.filename}: {e}", exc_info=True)
        print(f"ERRO NO PROCESSAMENTO: {e}")  # Print adicional para debug no terminal
        raise HTTPException(
            status_code=500,
            detail=f"Erro ao processar o arquivo: {str(e)}"
        )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=8000,
        reload=True
    )
