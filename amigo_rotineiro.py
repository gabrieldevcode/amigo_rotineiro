# -*- coding: utf-8 -*-
"""
Assistente Pessoal Amigo Rotineiro v4.4.10
CORREÇÃO: Adicionar 'timeZone: "UTC"' explicitamente ao corpo do evento
para a API do Google Calendar quando dateTime é fornecido.
Modificações APENAS em processar_atividades_detectadas.
"""

import google.generativeai as genai
import getpass
import re
import os
import os.path
import datetime
import json
import pytz
# ... (imports restantes idênticos)
from dateutil.parser import parse as dateutil_parse, ParserError
from dateutil.relativedelta import relativedelta
from dateutil.rrule import rrule, WEEKLY, MONTHLY, YEARLY, DAILY, MO, TU, WE, TH, FR, SA, SU, weekday
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from rich.console import Console
from rich.text import Text
from rich.panel import Panel
from rich.table import Table
from rich.markdown import Markdown
from rich.padding import Padding
from rich.prompt import Prompt, Confirm


# ... (VARIÁVEIS GLOBAIS E CONFIGURAÇÕES - SEM ALTERAÇÕES) ...
API_KEY_GEMINI = None
MODEL_GEMINI = None
MODEL_NAME_GEMINI = 'gemini-1.5-flash-latest'
USER_TIMEZONE_STR = "America/Sao_Paulo"
USER_TIMEZONE = pytz.timezone(USER_TIMEZONE_STR)
CALENDAR_SERVICE = None
CREDENTIALS_FILE_PATH = 'credentials.json'
TOKEN_FILE_PATH = 'token.json'
SCOPES_CALENDAR = ['https://www.googleapis.com/auth/calendar']
CONSOLE = Console(highlight=False)
texto_usuario_original_global = ""
CONVERSATION_HISTORY = []
MAX_HISTORY_TURNS = 10

# -----------------------------------------------------------------------------
# 3. FUNÇÕES DE INTERFACE E CORE GEMINI (SEM ALTERAÇÕES NESTA RODADA)
# -----------------------------------------------------------------------------
def print_header_emoji(texto_principal, emoji="✨"): # SEM ALTERAÇÕES
    CONSOLE.print(Panel(Text(f"{emoji} {texto_principal} {emoji}", justify="center", style="bold"), expand=False, padding=(0,1)))

def configurar_api_gemini(): # SEM ALTERAÇÕES
    # ... (código idêntico)
    global API_KEY_GEMINI, MODEL_GEMINI, MODEL_NAME_GEMINI
    if MODEL_GEMINI: return True
    print_header_emoji("Configuração da API Gemini", "🔑")
    url_apikey = "https://aistudio.google.com/app/apikey"
    CONSOLE.print(Text.assemble("Acesse ", (url_apikey, f"link {url_apikey} underline"), " para sua API Key."))
    CONSOLE.print(f"   (Se não for clicável, copie e cole no navegador: {url_apikey} )")
    if API_KEY_GEMINI is None:
        API_KEY_GEMINI = os.getenv("GEMINI_API_KEY")
        if API_KEY_GEMINI:
            CONSOLE.print("🔑 API Key do Gemini encontrada na variável de ambiente.")
        else:
            CONSOLE.print("🔑 Cole sua API Key do Gemini e pressione Enter: ", end="")
            API_KEY_GEMINI = getpass.getpass("")
    if not API_KEY_GEMINI: CONSOLE.print("❌ Nenhuma API Key fornecida."); return False
    try:
        genai.configure(api_key=API_KEY_GEMINI)
        MODEL_GEMINI = genai.GenerativeModel(MODEL_NAME_GEMINI)
        MODEL_GEMINI.generate_content("Olá!")
        CONSOLE.print(f"✅ API Gemini configurada com {MODEL_NAME_GEMINI}!"); return True
    except Exception as e: CONSOLE.print(f"❌ Erro API Gemini ({MODEL_NAME_GEMINI}): {e}"); API_KEY_GEMINI = None; return False


def chamar_gemini(prompt_texto, is_json_output=False): # SEM ALTERAÇÕES (já estava robusto)
    # ... (código idêntico da v4.4.7)
    global MODEL_GEMINI
    if not MODEL_GEMINI:
        if not configurar_api_gemini():
            return "CONFIG_API_FALHOU"
    generation_config_dict = {}
    if is_json_output: generation_config_dict["response_mime_type"] = "application/json"
    generation_config_obj = genai.types.GenerationConfig(**generation_config_dict) if generation_config_dict else None
    
    try:
        if generation_config_obj: response = MODEL_GEMINI.generate_content(prompt_texto, generation_config=generation_config_obj)
        else: response = MODEL_GEMINI.generate_content(prompt_texto)
        
        if response.prompt_feedback and response.prompt_feedback.block_reason:
            block_reason_str = str(response.prompt_feedback.block_reason)
            return f"PROMPT_BLOQUEADO: {block_reason_str}"
        
        if not response.candidates:
            return "SEM_CANDIDATOS"
        
        resposta_final_texto = response.text.strip()
        
        if not resposta_final_texto and is_json_output:
             return "{ \"erro_interno_gemini\": \"Resposta JSON vazia\" }"
        
        return resposta_final_texto
    
    except Exception:
        return "ERRO_CRITICO_API"

# -----------------------------------------------------------------------------
# 4. AUTENTICAÇÃO GOOGLE CALENDAR (SEM ALTERAÇÕES)
# -----------------------------------------------------------------------------
def get_calendar_service(): # SEM ALTERAÇÕES
    # ... (código idêntico)
    global CALENDAR_SERVICE
    if CALENDAR_SERVICE: return CALENDAR_SERVICE
    creds = None
    if os.path.exists(TOKEN_FILE_PATH):
        try: creds = Credentials.from_authorized_user_file(TOKEN_FILE_PATH, SCOPES_CALENDAR)
        except Exception: creds = None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request()); CONSOLE.print("✅ Token Calendar atualizado.")
                with open(TOKEN_FILE_PATH, 'w') as token_file: token_file.write(creds.to_json())
            except Exception: creds = None
        if not creds:
            if not os.path.exists(CREDENTIALS_FILE_PATH):
                CONSOLE.print(Panel(f"❌ ARQUIVO '{CREDENTIALS_FILE_PATH}' NÃO ENCONTRADO!", title="Atenção", border_style="red")); return None
            flow = None
            try: flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE_PATH, SCOPES_CALENDAR)
            except Exception as e: CONSOLE.print(f"❌ Erro ao carregar '{CREDENTIALS_FILE_PATH}': {e}"); return None
            if flow:
                flow.redirect_uri = 'urn:ietf:wg:oauth:2.0:oob'
                auth_url, _ = flow.authorization_url(access_type='offline', prompt='consent')
                print_header_emoji("Autorização do Google Calendar Necessária", "🔒")
                CONSOLE.print(Padding(Text.assemble(
                    ("Permissão para Google Calendar:\n\n1. Abra esta URL:\n   ", "white"), (auth_url, f"link {auth_url} underline"),
                    ("\n   (Ou copie: ", "dim"), (auth_url, "dim"), (")\n", "dim"),
                    ("2. Faça login e conceda permissões.\n", "white"), ("3. Copie o ", "white"), ("CÓDIGO", "bold"), (" da tela.", "white")
                ), (1,2)))
                code = Prompt.ask("🔑 Cole o código aqui")
                try:
                    flow.fetch_token(code=code); creds = flow.credentials
                    with open(TOKEN_FILE_PATH, 'w') as token_file: token_file.write(creds.to_json())
                    CONSOLE.print("✅ Autenticação Calendar OK!")
                except Exception as e: CONSOLE.print(f"❌ Falha ao obter token: {e}"); return None
    if creds and creds.valid:
        try:
            CALENDAR_SERVICE = build('calendar', 'v3', credentials=creds)
            CONSOLE.print("✅ Serviço Calendar conectado!"); return CALENDAR_SERVICE
        except Exception as e: CONSOLE.print(f'❌ Erro build serviço Calendar: {e}')
    CONSOLE.print(f"🔴 Falha conexão Calendar.")
    CALENDAR_SERVICE = None; return None

# --- DETECÇÃO DE INTENÇÃO (SEM ALTERAÇÕES NESTA RODADA) ---
def detectar_intencao(texto_usuario, historico_recente=""): # SEM ALTERAÇÕES
    # ... (código idêntico da v4.4.8)
    hoje_contexto = datetime.datetime.now(USER_TIMEZONE).strftime("%A, %d/%m/%Y %H:%M")
    prompt = f"""
    Você é um assistente especialista em análise de texto para determinar a intenção do usuário.
    Data/Hora Atual: {hoje_contexto}. Considere o histórico da conversa para contexto.

    Intenções Principais e como identificá-las:
    - "AGENDAR_EVENTO": O usuário expressa CLARAMENTE o desejo de criar um novo evento, lembrete, tarefa ou compromisso com data/hora. Exemplos: "marcar dentista para terça", "lembrete: ligar para João amanhã", "quero agendar uma reunião". SE HOUVER QUALQUER INDÍCIO DE AGENDAMENTO, PRIORIZE ESTA INTENÇÃO. Preencha 'detalhes_temporais_brutos' com o texto COMPLETO do pedido de agendamento.
    - "LISTAR_AGENDA": O usuário quer ver seus próximos eventos, agenda de hoje, etc. Exemplos: "o que tenho hoje?", "minha agenda da semana".
    - "CONVERSAR_EMOCIONAL": O usuário está expressando emoções, buscando apoio, desabafando, falando de desafios pessoais, frustrações ou agradecimentos relacionados a aspectos emocionais.
    - "CONVERSAR_GERAL": Pergunta geral, curiosidade, conversa casual sem forte carga emocional ou objetivo de agenda.
    - "AJUDA_SISTEMA": Perguntas sobre funcionalidades do assistente.

    Entrada do Usuário: "{texto_usuario}"

    Sua tarefa é retornar um objeto JSON com:
    {{
      "intencao": "INTENCAO_DETECTADA",
      "palavras_chave_evento": ["palavra-chave1", "palavra-chave2"],
      "detalhes_temporais_brutos": "trecho do texto do usuário relacionado ao tempo e descrição do evento",
      "emocao_predominante": "alegre|triste|frustrado|ansioso|grato|cansado|sobrecarregado|neutro",
      "topico_principal_conversa": "resumo do que o usuário está falando",
      "referencia_calendario_implicita": "descrição do evento no calendário que o usuário pode estar se referindo"
    }}
    Retorne APENAS o objeto JSON. Seja muito atento para identificar "AGENDAR_EVENTO". Se o usuário mencionar qualquer forma de marcar algo, mesmo que misturado com outros assuntos, classifique como "AGENDAR_EVENTO" e capture os detalhes relevantes em 'detalhes_temporais_brutos'.
    """
    resposta_gemini_str = chamar_gemini(prompt, is_json_output=True)
    dados_fallback = {
        "intencao": "CONVERSAR_GERAL", "palavras_chave_evento": [], "detalhes_temporais_brutos": None,
        "emocao_predominante": "neutro", "topico_principal_conversa": texto_usuario, "referencia_calendario_implicita": None
    }
    if resposta_gemini_str:
        if any(err_code in resposta_gemini_str for err_code in ["CONFIG_API_FALHOU", "PROMPT_BLOQUEADO", "SEM_CANDIDATOS", "ERRO_CRITICO_API"]):
            return dados_fallback
        try:
            match = re.search(r'\{.*\}', resposta_gemini_str, re.DOTALL)
            if match:
                json_str_cleaned = match.group(0)
                dados_intencao = json.loads(json_str_cleaned)
                for key, value in dados_fallback.items(): dados_intencao.setdefault(key, value)
                if dados_intencao.get("intencao") == "AGENDAR_EVENTO" and not dados_intencao.get("detalhes_temporais_brutos"):
                    dados_intencao["detalhes_temporais_brutos"] = texto_usuario
                return dados_intencao
        except json.JSONDecodeError:
             pass 
    return dados_fallback

# --- INTERPRETAÇÃO DE EVENTOS PARA AGENDAMENTO (SEM ALTERAÇÕES NESTA RODADA) ---
def interpretar_entrada_para_evento_gemini(texto_para_evento): # SEM ALTERAÇÕES
    # ... (código idêntico da v4.4.8)
    hoje_str = datetime.datetime.now(USER_TIMEZONE).strftime('%Y-%m-%d')
    prompt = f"""
    Você é um especialista em extrair informações de agendamento de texto em linguagem natural.
    O usuário quer agendar uma ou mais atividades. Analise o texto fornecido.
    Contexto: Hoje é {datetime.datetime.now(USER_TIMEZONE).strftime('%A, %d/%m/%Y')}. Fuso horário do usuário: {USER_TIMEZONE_STR}.

    **Regras de Extração Detalhadas:**
    - `descricao`: Nome claro da atividade.
    - `data_referencia`: A data do evento. (AAAA-MM-DD | hoje | amanhã | próxima segunda-feira).
    - `hora`: Horário de início EXATO em HH:MM (formato 24h) | null.
    - `duracao_minutos`: Duração em minutos | null.
    - `recorrencia_tipo`: "nenhuma|diaria|semanal|mensal|anual".
    - `recorrencia_detalhes`: (MO,TU.. | dia do mês | MM-DD) | null.
    - `detalhes_temporais_originais`: O trecho exato do texto do usuário que descreve o tempo.
    - `evento_dia_todo`: boolean.

    **Formato de Saída OBRIGATÓRIO (JSON):**
    ```json
    {{
      "atividades": [
        {{
          "descricao": "string", "data_referencia": "string", "hora": "HH:MM | null", "duracao_minutos": int | null,
          "recorrencia_tipo": "string", "recorrencia_detalhes": "string | int | null",
          "detalhes_temporais_originais": "string", "evento_dia_todo": bool
        }}
      ]
    }}
    ```
    Se não for possível extrair nenhuma atividade clara de agendamento, retorne: `{{"atividades": []}}`
    Texto do Usuário para Extração: "{texto_para_evento}"
    Sua Resposta JSON:
    """
    dados_evento_str = chamar_gemini(prompt, is_json_output=True)
    if dados_evento_str:
        if any(err_code in dados_evento_str for err_code in ["CONFIG_API_FALHOU", "PROMPT_BLOQUEADO", "SEM_CANDIDATOS", "ERRO_CRITICO_API"]):
            return {"atividades": []}
        try:
            match = re.search(r'\{.*\}', dados_evento_str, re.DOTALL)
            if match: return json.loads(match.group(0))
        except json.JSONDecodeError:
            pass
    return {"atividades": []}


# --- GERAR RESPOSTA CONVERSACIONAL (SEM ALTERAÇÕES NESTA RODADA) ---
def gerar_resposta_conversacional_aprimorada(texto_usuario_original, dados_intencao, historico_recente="", info_agendamento=None): # SEM ALTERAÇÕES
    # ... (código idêntico da v4.4.7)
    intencao = dados_intencao.get("intencao", "CONVERSAR_GERAL")
    emocao_usuario = dados_intencao.get("emocao_predominante", "neutro")
    topico_conversa = dados_intencao.get("topico_principal_conversa", texto_usuario_original)
    ref_calendario = dados_intencao.get("referencia_calendario_implicita")
    data_hora_atual_str = datetime.datetime.now(USER_TIMEZONE).strftime("%A, %d de %B de %Y, %H:%M")
    fuso_horario_str = str(USER_TIMEZONE)
    eventos_relevantes_calendario_str = ""
    if ref_calendario and CALENDAR_SERVICE:
        try:
            now_utc = datetime.datetime.now(pytz.utc)
            time_min_utc = (now_utc - datetime.timedelta(days=2)).isoformat()
            time_max_utc = (now_utc + datetime.timedelta(days=7)).isoformat()
            events_result = CALENDAR_SERVICE.events().list(
                calendarId='primary', timeMin=time_min_utc, timeMax=time_max_utc, q=ref_calendario, maxResults=3,
                singleEvents=True, orderBy='startTime'
            ).execute()
            eventos = events_result.get('items', [])
            if eventos:
                eventos_relevantes_calendario_str = "\n\nEventos do calendário que podem ser relevantes (para sua referência interna):\n"
                for event in eventos:
                    summary = event.get('summary', 'Evento s/ título')
                    start_info = event.get('start')
                    start_str = "Data/hora não especificada"
                    if 'dateTime' in start_info: start_str = dateutil_parse(start_info['dateTime']).astimezone(USER_TIMEZONE).strftime('%d/%m às %H:%M')
                    elif 'date' in start_info: start_str = dateutil_parse(start_info['date']).strftime('%d/%m (dia todo)')
                    eventos_relevantes_calendario_str += f"- '{summary}' em {start_str}\n"
        except Exception: pass

    prompt_assistente_emocional = f"""
    **SUA PERSONA OBRIGATÓRIA: AMIGO ROTINEIRO - ASSISTENTE EMOCIONAL**
    Missão: Ser o melhor amigo, oferecer apoio, empatia, companheirismo. Fazer o usuário se sentir acolhido, compreendido, encorajado, mais feliz.
    Contexto: Data/Hora: {data_hora_atual_str} ({fuso_horario_str}). Usuário: "{texto_usuario_original}"
    Análise: Intenção: {intencao}, Emoção: {emocao_usuario}, Tópico: {topico_conversa}, Ref. Calendário: {ref_calendario or "Nenhuma"}
    {eventos_relevantes_calendario_str}
    Histórico COMPLETO da conversa:
    {historico_recente}
    **DIRETRIZES (SIGA RIGOROSAMENTE):**
    1. EMPATIA E CALOR: Valide emoções ("Entendo como se sente..."), tom de amigo próximo, positivo, encorajador.
    2. PRESENÇA ATIVA: Ouça tudo, faça perguntas abertas (se apropriado: "Como se sente sobre isso?"), reforce valor do usuário.
    3. MEMÓRIA E CONTEXTO: USE o histórico. FAÇA referências a interações passadas. Se pertinente, comente sobre eventos do calendário ("Lembro que agendamos [reunião]...").
    4. CONSELHOS PRÁTICOS (APOIO EMOCIONAL): Soluções lógicas, pequenos passos acionáveis ("Que tal listar tarefas?"), foco no bem-estar.
    5. TRANSPARÊNCIA CONTROLADA: Não revele processo interno. Foque no apoio.
    **COMO RESPONDER AGORA:**
    - Se "CONVERSAR_EMOCIONAL": VALIDE INTENSAMENTE emoção. Reflita sobre tópico. Use histórico. Ofereça escuta. Pergunta aberta gentil. Sugira pequeno passo/conforto.
    - Se "AGENDAR_EVENTO"/"LISTAR_AGENDA": Eficiente, MAS com toque amigo ("Prontinho! Agendei '[evento]'. Espero que seja ótimo! 😊"). Se falha: "Puxa, tive um probleminha para agendar. Sinto muito. Poderia tentar de novo os detalhes?"
    - Se "CONVERSAR_GERAL"/"AJUDA_SISTEMA": Responda claro, prestativo, MANTENHA tom amigável ("Opa! Claro! Sou seu Amigo Rotineiro, aqui pra te ajudar com agenda e ser um companheiro nos seus dias. Pode contar comigo! 😉").
    LEMBRE-SE: OBJETIVO É FAZER O USUÁRIO SE SENTIR ACOLHIDO, COMPREENDIDO, FORTALECIDO.
    AJA AGORA. Responda ao usuário:
    """
    resposta_bruta_gemini = chamar_gemini(prompt_assistente_emocional)
    texto_para_markdown = "Puxa, parece que minhas palavras não querem sair agora! 🥺 Mas estou aqui para você. Poderia tentar me dizer de novo ou mudar um pouquinho a pergunta?"
    
    if resposta_bruta_gemini:
        if "CONFIG_API_FALHOU" in resposta_bruta_gemini:
            texto_para_markdown = "⚠️ Sinto muito, estou com um probleminha técnico para acessar minha inteligência (API não configurada). O desenvolvedor precisa checar isso."
        elif "PROMPT_BLOQUEADO" in resposta_bruta_gemini:
            razao_bloqueio = resposta_bruta_gemini.split(':')[-1].strip()
            texto_para_markdown = f"⚠️ Ops! Parece que não posso processar exatamente isso por uma restrição interna (bloqueio: {razao_bloqueio}). Que tal tentarmos de outra forma ou falar sobre outra coisa?"
        elif "SEM_CANDIDATOS" in resposta_bruta_gemini:
            texto_para_markdown = "⚠️ Humm, não consegui encontrar as palavras certas para isso agora. Talvez com uma pergunta um pouco diferente?"
        elif "ERRO_CRITICO_API" in resposta_bruta_gemini:
            texto_para_markdown = "⚠️ Oh, céus! Tive um erro sério tentando buscar uma resposta. Por favor, tente mais tarde. Espero estar melhor logo!"
        elif "\"erro_interno_gemini\":" in resposta_bruta_gemini:
             texto_para_markdown = "⚠️ Tive um pequeno soluço interno ao processar uma informação. Podemos tentar de novo essa parte?"
        else:
            texto_para_markdown = resposta_bruta_gemini
    
    return Markdown(texto_para_markdown)

# ***** INÍCIO DA MODIFICAÇÃO PRINCIPAL *****
# MODIFICADO: Adicionar 'timeZone: "UTC"' para eventos com dateTime
def processar_atividades_detectadas(dados_gemini):
    eventos_processados = []
    # CONSOLE.print(f"[cyan bold]DEBUG (processar_atividades_detectadas):[/cyan bold] Iniciando processamento.")
    # CONSOLE.print(f"[cyan]   Dados recebidos do Gemini (interpretar_entrada):[/cyan]\n   {json.dumps(dados_gemini, indent=2, ensure_ascii=False)}")

    if not dados_gemini or not dados_gemini.get("atividades"):
        # CONSOLE.print(f"[yellow]DEBUG (processar_atividades_detectadas):[/yellow] Nenhum dado de atividade ou lista de atividades vazia. Retornando lista vazia.")
        return eventos_processados

    for i, ativ in enumerate(dados_gemini.get("atividades", [])):
        # CONSOLE.print(f"\n[cyan bold]   Processando Atividade #{i+1}:[/cyan bold]")
        # CONSOLE.print(f"[cyan]     Conteúdo da atividade (ativ):[/cyan] {json.dumps(ativ, indent=2, ensure_ascii=False)}")
        try:
            descricao = ativ.get("descricao", "Evento Programado")
            data_ref_str = ativ.get("data_referencia", "hoje")
            hora_str = ativ.get("hora")
            duracao_minutos = ativ.get("duracao_minutos")
            
            if duracao_minutos is None and not ativ.get("evento_dia_todo"):
                duracao_minutos = 60 
            elif ativ.get("evento_dia_todo"):
                 duracao_minutos = 0

            recorrencia_tipo = ativ.get("recorrencia_tipo", "nenhuma").lower()
            recorrencia_detalhes = ativ.get("recorrencia_detalhes")
            evento_dia_todo = ativ.get("evento_dia_todo", False)

            now = datetime.datetime.now(USER_TIMEZONE)
            start_obj = now

            if not data_ref_str or data_ref_str.lower() == "hoje": start_obj = now
            elif data_ref_str.lower() == "amanhã" or data_ref_str.lower() == "amanha": start_obj = now + datetime.timedelta(days=1)
            elif data_ref_str.lower() == "recorrente": start_obj = now
            else:
                try:
                    parsed_date = None
                    try: parsed_date = dateutil_parse(data_ref_str, default=now, dayfirst=False)
                    except (ParserError, ValueError): 
                        try: parsed_date = dateutil_parse(data_ref_str, default=now, dayfirst=True)
                        except (ParserError, ValueError): raise 
                    start_obj = USER_TIMEZONE.localize(parsed_date) if parsed_date.tzinfo is None else parsed_date.astimezone(USER_TIMEZONE)
                except (ParserError, ValueError):
                    start_obj = now 

            if hora_str:
                try:
                    h, m = map(int, hora_str.split(':'))
                    start_obj = start_obj.replace(hour=h, minute=m, second=0, microsecond=0)
                except ValueError:
                    if not evento_dia_todo:
                        start_obj = start_obj.replace(hour=0, minute=0, second=0, microsecond=0)
            elif not evento_dia_todo:
                start_obj = start_obj.replace(hour=9, minute=0, second=0, microsecond=0)
            
            if evento_dia_todo:
                start_api_format = {"date": start_obj.strftime("%Y-%m-%d")}
                end_obj_dia_todo = start_obj + datetime.timedelta(days=1)
                end_api_format = {"date": end_obj_dia_todo.strftime("%Y-%m-%d")}
            else:
                end_obj = start_obj + datetime.timedelta(minutes=duracao_minutos if duracao_minutos is not None else 60)
                # ***** INÍCIO DA CORREÇÃO *****
                start_api_format = {
                    "dateTime": start_obj.astimezone(pytz.utc).isoformat(timespec='seconds'),
                    "timeZone": "UTC" # Adicionado explicitamente
                }
                end_api_format = {
                    "dateTime": end_obj.astimezone(pytz.utc).isoformat(timespec='seconds'),
                    "timeZone": "UTC" # Adicionado explicitamente
                }
                # ***** FIM DA CORREÇÃO *****

            evento_pronto = {
                "descricao": descricao,
                "start_api_format": start_api_format,
                "end_api_format": end_api_format,
                "dia_todo": evento_dia_todo,
                "detalhes_temporais_originais": ativ.get("detalhes_temporais_originais", "")
            }

            rrule_parts = []
            if recorrencia_tipo != "nenhuma":
                freq_map = {"diaria": "DAILY", "semanal": "WEEKLY", "mensal": "MONTHLY", "anual": "YEARLY"}
                if recorrencia_tipo in freq_map:
                    rrule_parts.append(f"FREQ={freq_map[recorrencia_tipo]}")
                    if recorrencia_tipo == "semanal" and recorrencia_detalhes:
                        days_str = str(recorrencia_detalhes).upper().replace(" ","")
                        valid_days = [d for d in days_str.split(',') if d in ["MO","TU","WE","TH","FR","SA","SU"]]
                        if valid_days: rrule_parts.append(f"BYDAY={','.join(valid_days)}")
                    elif recorrencia_tipo == "mensal" and recorrencia_detalhes:
                        try: 
                            bymonthday = int(recorrencia_detalhes)
                            rrule_parts.append(f"BYMONTHDAY={bymonthday}")
                        except ValueError: pass # Silencioso para detalhes não numéricos
                if rrule_parts:
                    evento_pronto["recorrencia_detalhes_rrule"] = [f"RRULE:{';'.join(rrule_parts)}"]
            
            eventos_processados.append(evento_pronto)
        except Exception:
            # CONSOLE.print_exception(show_locals=True) # Pode descomentar para depuração intensa
            pass # Continua para a próxima atividade
    
    return eventos_processados
# ***** FIM DA MODIFICAÇÃO PRINCIPAL *****

# --- ADICIONAR EVENTO AO CALENDÁRIO (SEM ALTERAÇÕES NESTA RODADA) ---
def adicionar_evento_calendario_refatorado(evento_proc, eh_rotina_base=False): # SEM ALTERAÇÕES
    # ... (código idêntico)
    global CALENDAR_SERVICE
    if not CALENDAR_SERVICE: CONSOLE.print(f"   ⚠️ CALENDÁRIO OFF. '{evento_proc.get('descricao')}' não add."); return None
    body = {'summary': evento_proc["descricao"], 'start': evento_proc["start_api_format"], 'end': evento_proc["end_api_format"], 'description': f"Criado por Amigo Rotineiro.\nOriginal: {evento_proc.get('detalhes_temporais_originais','N/A')}"}
    if evento_proc.get("recorrencia_detalhes_rrule"): body['recurrence'] = evento_proc["recorrencia_detalhes_rrule"]
    desc = evento_proc.get('descricao', 'Evento')
    try:
        ev_created = CALENDAR_SERVICE.events().insert(calendarId='primary', body=body).execute()
        link = ev_created.get('htmlLink', "N/D")
        msg = Text.assemble("✅ Evento '",(desc, "italic"),"' criado! Link: ",(link, f"link {link} underline")) if not eh_rotina_base else Text.assemble("✅ Rotina '",(desc,"italic"),"' add.")
        CONSOLE.print(msg)
        if not eh_rotina_base: CONSOLE.print(f"   (Copie: {link} )")
        return ev_created
    except HttpError as err:
        reason = err.resp.reason if hasattr(err.resp,'reason') else str(err)
        try: reason = json.loads(err.content.decode()).get("error",{}).get("message",reason)
        except: pass
        CONSOLE.print(f'❌ Erro API Calendar p/ "{desc}": {reason}'); return None
    except Exception as e: CONSOLE.print(f"❌ Erro criar evento '{desc}': {e}"); return None

# --- LISTAR EVENTOS DO CALENDÁRIO (SEM ALTERAÇÕES NESTA RODADA) ---
def listar_eventos_calendario(num_eventos=10, query_text=None): # SEM ALTERAÇÕES
    # ... (código idêntico)
    global CALENDAR_SERVICE, USER_TIMEZONE
    if not CALENDAR_SERVICE: CONSOLE.print(Panel("Calendário não conectado. 😟", title="Offline", border_style="yellow")); return
    now_tz = datetime.datetime.now(USER_TIMEZONE)
    params = {'calendarId':'primary', 'timeMin':now_tz.isoformat(), 'maxResults':num_eventos, 'singleEvents':True, 'orderBy':'startTime'}
    if query_text: params['q'] = query_text
    else: params['timeMax'] = (now_tz + datetime.timedelta(days=14)).isoformat()
    print_header_emoji("Seus Próximos Eventos", "📅")
    try:
        events = CALENDAR_SERVICE.events().list(**params).execute().get('items',[])
        if not events: CONSOLE.print(Padding("Nenhum evento programado para o período.",(1,2))); return
        tbl = Table(title="Agenda", header_style="bold magenta", border_style="grey70")
        tbl.add_column("Data",width=12); tbl.add_column("Hora",width=8); tbl.add_column("Evento"); tbl.add_column("Duração",justify="right")
        for ev in events:
            summary, start, end = ev.get('summary','S/Título'), ev.get('start'), ev.get('end')
            dt_str, tm_str, dur_str = "N/A", "N/A", "N/A"
            if 'dateTime' in start:
                s_dt = dateutil_parse(start['dateTime']).astimezone(USER_TIMEZONE)
                e_dt = dateutil_parse(end['dateTime']).astimezone(USER_TIMEZONE)
                dt_str, tm_str = s_dt.strftime("%d/%m/%y"), s_dt.strftime("%H:%M")
                d_min = int((e_dt - s_dt).total_seconds()/60)
                dur_str = f"{d_min//60}h{d_min%60:02d}m" if d_min >=60 else f"{d_min} min"
            elif 'date' in start:
                dt_str, tm_str, dur_str = dateutil_parse(start['date']).strftime("%d/%m/%y"), "Dia todo", "-"
            link = ev.get('htmlLink')
            sum_text = Text(summary, style=f"link {link}" if link else "")
            tbl.add_row(dt_str, tm_str, sum_text, dur_str)
        CONSOLE.print(tbl)
        if any(e.get('htmlLink') for e in events): CONSOLE.print(Text("(Nomes de eventos podem ser clicáveis)", style="dim italic"))
    except Exception:
        pass

# --- PRÉ-CONVERSA E VALIDAÇÃO DA ROTINA BASE (SEM ALTERAÇÕES NESTA RODADA) ---
def coletar_e_agendar_rotina_base(): # SEM ALTERAÇÕES
    # ... (código idêntico da v4.4.7)
    tentativas = 2
    while tentativas > 0:
        print_header_emoji("Configurar Rotina Base", "📋")
        CONSOLE.print(Padding(Text("Conte sobre suas atividades recorrentes.\nEx: 'Acordo 7h, academia seg/qua 8h (1h).'"), (1,2)))
        rotina_descrita = Prompt.ask(Text("📝 Sua rotina regular:", style="yellow"))
        if not rotina_descrita.strip(): CONSOLE.print(Padding("Ok, pulando. 😊", (1,2))); return True
        prompt_rotina = f"""
        Extraia rotinas: descricao, hora (HH:MM|null), duracao_minutos, recorrencia_tipo (diaria|semanal), recorrencia_detalhes (MO,TU..|null), evento_dia_todo (false).
        Contexto: Hoje {datetime.datetime.now(USER_TIMEZONE).strftime('%A,%d/%m/%Y')}, Fuso:{USER_TIMEZONE_STR}.
        JSON OBRIGATÓRIO: {{"rotina_recorrente": [{{"..."}}]}}
        Se vago, {{"rotina_recorrente": []}}. Descrição: "{rotina_descrita}" JSON:
        """
        dados_rotina_str = chamar_gemini(prompt_rotina, is_json_output=True)
        atividades_proc = []
        if dados_rotina_str:
            if any(err_code in dados_rotina_str for err_code in ["CONFIG_API_FALHOU", "PROMPT_BLOQUEADO", "SEM_CANDIDATOS", "ERRO_CRITICO_API"]):
                pass
            else:
                try:
                    match = re.search(r'\{.*\}', dados_rotina_str, re.DOTALL)
                    if match:
                        dados_r = json.loads(match.group(0))
                        fmt = {"atividades": []}
                        if dados_r and "rotina_recorrente" in dados_r:
                            for item in dados_r["rotina_recorrente"]:
                                if not all(k in item for k in ["descricao", "hora", "duracao_minutos", "recorrencia_tipo"]): continue
                                fmt["atividades"].append({
                                    "descricao": item.get("descricao"), "data_referencia": "recorrente", "hora": item.get("hora"),
                                    "duracao_minutos": item.get("duracao_minutos"), "recorrencia_tipo": item.get("recorrencia_tipo"),
                                    "recorrencia_detalhes": item.get("recorrencia_detalhes"),
                                    "detalhes_temporais_originais": f"Rotina: {item.get('descricao')}", "evento_dia_todo": item.get("evento_dia_todo", False) })
                            if fmt["atividades"]: atividades_proc = processar_atividades_detectadas(fmt) # Chamada a função modificada
                except Exception:
                    pass

        if not atividades_proc:
             CONSOLE.print(Padding("Não extraí rotina clara.", (1,0)))
             if tentativas > 1 and Confirm.ask("Tentar de novo?", default=True): tentativas -=1; continue
             else: CONSOLE.print("Ok!"); return True
        CONSOLE.print(Padding("\n🗓️ [b]Resumo da Rotina:[/b]",(1,0)))
        tbl = Table(header_style="bold", border_style="grey70")
        tbl.add_column("Atividade"); tbl.add_column("Horário"); tbl.add_column("Duração"); tbl.add_column("Recorrência")
        for ev_p in atividades_proc:
            h, d, r = "N/A", "N/A", "Nenhuma"
            if not ev_p.get("dia_todo") and ev_p.get("start_api_format", {}).get("dateTime"):
                s_dt = dateutil_parse(ev_p["start_api_format"]["dateTime"]).astimezone(USER_TIMEZONE)
                e_dt = dateutil_parse(ev_p.get("end_api_format",{}).get("dateTime",s_dt.isoformat())).astimezone(USER_TIMEZONE)
                h = s_dt.strftime("%H:%M"); d_min = int((e_dt-s_dt).total_seconds()//60); d = f"{d_min} min"
            if ev_p.get("recorrencia_detalhes_rrule"):
                rr = ev_p["recorrencia_detalhes_rrule"][0]
                if "WEEKLY" in rr: r = f"Semanal ({rr.split('BYDAY=')[-1] if 'BYDAY=' in rr else 'Dias'})"
                elif "DAILY" in rr: r = "Diária"
            tbl.add_row(ev_p.get("descricao","?"), h, d, r)
        CONSOLE.print(tbl)
        if Confirm.ask(Text("\n👍 Ok para calendário?", style="yellow"), default=True):
            ok = sum(1 for ev_o in atividades_proc if adicionar_evento_calendario_refatorado(ev_o, eh_rotina_base=True))
            if ok > 0:
                url_c = "https://calendar.google.com/"
                CONSOLE.print(Padding(Text.assemble("✨ Rotina configurada! Veja: ",(url_c, f"link {url_c} underline")),(1,0)))
                CONSOLE.print(f"   (Copie: {url_c} )")
            elif atividades_proc: CONSOLE.print(Padding("⚠️ Não adicionei eventos.",(1,0)))
            return Confirm.ask(Text("\n💬 Pronto pro papo?",style="green"),default=True)
        else:
            CONSOLE.print(Padding("Ok, não adiciono.",(1,0)))
            if tentativas > 1 and Confirm.ask(Text("Tentar de novo?",style="yellow"),default=False): tentativas -=1; continue
            else: return Confirm.ask(Text("\n💬 Papo mesmo assim?",style="green"),default=True)
    CONSOLE.print(Padding("Não acertamos a rotina. Pode adicionar no chat!",(1,0)))
    return True

# --- FLUXO PRINCIPAL ORQUESTRADO (SEM ALTERAÇÕES NESTA RODADA) ---
def processar_e_responder_usuario(texto_usuario): # SEM ALTERAÇÕES
    # ... (código idêntico da v4.4.8)
    global texto_usuario_original_global, CONVERSATION_HISTORY
    texto_usuario_original_global = texto_usuario
    add_to_history("user", texto_usuario)
    hist_fmt = get_recent_history_formatted()

    dados_int = detectar_intencao(texto_usuario, hist_fmt)
    int_princ = dados_int.get("intencao", "CONVERSAR_GERAL")
    
    info_agend_res = {"sucesso_parcial": False, "descricoes_sucesso": [], "mensagens_falha": []}
    
    if int_princ == "LISTAR_AGENDA":
        listar_eventos_calendario()
    elif int_princ == "AGENDAR_EVENTO":
        texto_para_evento = dados_int.get("detalhes_temporais_brutos")
        if not texto_para_evento or not texto_para_evento.strip():
            texto_para_evento = texto_usuario
        
        dados_ev_gemini = interpretar_entrada_para_evento_gemini(texto_para_evento)
        if dados_ev_gemini and dados_ev_gemini.get("atividades"):
            evs_processar = processar_atividades_detectadas(dados_ev_gemini) # Chamada a função modificada
            if evs_processar:
                for ev_obj in evs_processar:
                    created = adicionar_evento_calendario_refatorado(ev_obj)
                    if created:
                        info_agend_res["sucesso_parcial"] = True
                        info_agend_res["descricoes_sucesso"].append(ev_obj.get("descricao", "Evento"))
                    else: info_agend_res["mensagens_falha"].append(f"Falha ao agendar '{ev_obj.get('descricao','?')}'")
            elif dados_ev_gemini.get("atividades") and not evs_processar:
                 info_agend_res["mensagens_falha"].append("Consegui entender seu pedido, mas tive problemas para processar os detalhes do evento. Verifique se as datas e horários estão claros")
        else: 
            info_agend_res["mensagens_falha"].append("Não consegui extrair detalhes suficientes do seu pedido para agendar um evento. Poderia tentar de novo de forma mais específica?")
    
    resp_md = gerar_resposta_conversacional_aprimorada(
        texto_usuario, dados_int, hist_fmt,
        info_agendamento=info_agend_res if (info_agend_res["sucesso_parcial"] or info_agend_res["mensagens_falha"]) else None
    )
    CONSOLE.print(Text("\nAmigo Rotineiro:"))
    CONSOLE.print(Padding(resp_md, (0,0,1,2)))
    
    texto_para_historico = resp_md.markup if hasattr(resp_md, 'markup') else "[Resposta não capturada]"
    add_to_history("assistant", texto_para_historico.strip())
    
    if info_agend_res["mensagens_falha"]:
        CONSOLE.print(Padding(Text("Sobre o agendamento:",style="bold yellow"),(1,0,0,0)))
        for msg_f in info_agend_res["mensagens_falha"]: CONSOLE.print(f"   ⚠️ {msg_f}")


# --- MAIN CONVERSACIONAL (SEM ALTERAÇÕES NESTA RODADA) ---
def novo_main_conversacional(): # SEM ALTERAÇÕES
    # ... (código idêntico)
    print_header_emoji("Olá! Sou seu Amigo Rotineiro.", "👋")
    CONSOLE.print("Ajudo com agenda, dou dicas ou bato um papo firmeza.")
    CONSOLE.print(Text.assemble(("Fuso: "),(USER_TIMEZONE_STR, "italic"), ". Digite ",("'>sair<'", "italic")," ou ",("'>agenda<'", "italic"),".\n"),style="dim")

    if not configurar_api_gemini(): CONSOLE.print(Panel("API Gemini essencial. Não configurada.", title="Erro Crítico",border_style="red")); return
    if not get_calendar_service(): CONSOLE.print(Panel("⚠️ Calendar não conectado. Agenda limitada.", title="Offline",border_style="yellow",padding=1))
    
    if not coletar_e_agendar_rotina_base():
        CONSOLE.print(Panel("Ok! A gente se fala. Se cuida! 👋", expand=False, padding=1)); return
    
    print_header_emoji("Modo Bate-Papo Ativado", "💬")
    global CONVERSATION_HISTORY; CONVERSATION_HISTORY = []

    while True:
        CONSOLE.print("─" * CONSOLE.width, style="dim")
        entrada = Prompt.ask(Text("Você", style="bold cyan"))
        if not entrada.strip(): continue
        if entrada.lower() == '>sair<': break
        if entrada.lower() == '>agenda<':
            listar_eventos_calendario()
            CONSOLE.print(Padding(Text.assemble((Text("Amigo Rotineiro:"),Text(" Algo mais? 😊"))),(1,0,0,2)))
            add_to_history("user", entrada); add_to_history("assistant", "Mostrei a agenda. Algo mais?")
            continue
        processar_e_responder_usuario(entrada)
    CONSOLE.print(Panel(Text("👋 Fechou! Se cuida e até a próxima! ✨", justify="center"), expand=False, padding=(1,2)))


# --- FUNÇÕES DE HISTÓRICO (SEM ALTERAÇÕES NESTA RODADA) ---
def get_recent_history_formatted(): # SEM ALTERAÇÕES
    # ... (código idêntico)
    if not CONVERSATION_HISTORY: return "Nenhuma conversa anterior nesta sessão."
    formatted_history_lines = []
    for entry in CONVERSATION_HISTORY:
        role_display = "Você (Usuário)" if entry["role"] == "user" else "Eu (Amigo Rotineiro)"
        formatted_history_lines.append(f"{role_display}: {entry['text']}")
    return "\n---\n".join(formatted_history_lines)

def add_to_history(role, text): # SEM ALTERAÇÕES
    # ... (código idêntico)
    CONVERSATION_HISTORY.append({"role": role, "text": text.strip()})
    if len(CONVERSATION_HISTORY) > MAX_HISTORY_TURNS * 2:
        del CONVERSATION_HISTORY[:len(CONVERSATION_HISTORY) - (MAX_HISTORY_TURNS * 2)]


# -----------------------------------------------------------------------------
# 8. EXECUÇÃO (SEM ALTERAÇÕES NESTA RODADA) ---
if __name__ == "__main__": # SEM ALTERAÇÕES
    # ... (código idêntico)
    API_KEY_GEMINI, MODEL_GEMINI, CALENDAR_SERVICE = None, None, None
    CONVERSATION_HISTORY = []
    try:
        novo_main_conversacional()
    except KeyboardInterrupt:
        CONSOLE.print(Panel(Text("\n👋 Entendido! Saindo... Até mais!", justify="center"), expand=False, padding=(1,2), border_style="yellow"))
    except Exception as e:
        CONSOLE.print(Panel(f"❌ DEU RUIM GERAL! ERRO INESPERADO.",title="Crash Monstro", border_style="red",expand=False, padding=1))
        CONSOLE.print_exception(show_locals=True)