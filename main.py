import os
import tempfile
import traceback
from pathlib import Path

from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

from gerarrelatorios import (
    MESES_PT,
    DadosAgregados,
    carregar_dados,
    gerar_excel_mensal,
    pdf_acumulado,
    pdf_anual,
    pdf_metas,
    pdf_mensal,
    pdf_regional,
    pdf_sem_regional,
)

app = FastAPI(
    title="DER-MG · Relatórios de Educação para o Trânsito",
    version="2.0.0",
)

if os.getenv("VERCEL"):
    OUTPUT_DIR = Path("/tmp/output")
else:
    BASE_DIR = Path(__file__).parent
    OUTPUT_DIR = BASE_DIR / "output"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def _gerar_relatorio(tipo: str, dados: DadosAgregados, ano: int, mes: int, out: str) -> str:
    dispatch = {
        "mensal":        lambda: pdf_mensal(dados, ano, mes, out),
        "acumulado":     lambda: pdf_acumulado(dados, ano, mes, out),
        "anual":         lambda: pdf_anual(dados, ano, out),
        "metas":         lambda: pdf_metas(dados, ano, out),
        "regional":      lambda: pdf_regional(dados, ano, out),
        "sem_regional":  lambda: pdf_sem_regional(dados, ano, out),
        "excel":         lambda: gerar_excel_mensal(dados, ano, mes, out),
    }
    if tipo not in dispatch:
        raise ValueError(f"Tipo de relatório desconhecido: {tipo}")
    return dispatch[tipo]()


HTML_PAGE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>DER-MG · Relatórios</title>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap" rel="stylesheet"/>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --red:      #CC0000;
    --red-dk:   #A80000;
    --bg:       #F5F5F5;
    --white:    #FFFFFF;
    --border:   #E0E0E0;
    --text:     #1A1A1A;
    --muted:    #757575;
    --label:    #555555;
    --success:  #2E7D32;
    --radius:   8px;
  }

  html, body {
    min-height: 100%;
    background: var(--bg);
    color: var(--text);
    font-family: 'Inter', sans-serif;
    font-size: 14px;
  }

  .wrapper {
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 40px 16px 60px;
  }

  /* ── Header ── */
  .header {
    width: 100%;
    max-width: 480px;
    margin-bottom: 28px;
    display: flex;
    flex-direction: column;
    align-items: flex-start;
    gap: 4px;
  }

  .header-badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: .09em;
    text-transform: uppercase;
    color: var(--red);
    margin-bottom: 2px;
  }

  .header-badge::before {
    content: '';
    display: inline-block;
    width: 6px;
    height: 6px;
    background: var(--red);
    border-radius: 50%;
  }

  .header h1 {
    font-size: 26px;
    font-weight: 600;
    color: var(--text);
    line-height: 1.2;
  }

  .header p {
    color: var(--muted);
    font-size: 13px;
    margin-top: 2px;
  }

  /* ── Card ── */
  .card {
    width: 100%;
    max-width: 480px;
    background: var(--white);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 28px 28px 24px;
    box-shadow: 0 2px 8px rgba(0,0,0,.07);
  }

  /* ── Fields ── */
  .fields {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
    margin-bottom: 16px;
  }

  .field { display: flex; flex-direction: column; gap: 6px; }
  .field.full { grid-column: 1 / -1; }

  label {
    font-size: 11px;
    font-weight: 600;
    letter-spacing: .07em;
    text-transform: uppercase;
    color: var(--label);
  }

  select {
    appearance: none;
    -webkit-appearance: none;
    background: var(--white);
    border: 1px solid var(--border);
    border-radius: 6px;
    color: var(--text);
    font-family: 'Inter', sans-serif;
    font-size: 14px;
    padding: 9px 34px 9px 12px;
    cursor: pointer;
    outline: none;
    transition: border-color .15s, box-shadow .15s;
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' fill='none'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%23999' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E");
    background-repeat: no-repeat;
    background-position: right 11px center;
  }

  select:focus {
    border-color: var(--red);
    box-shadow: 0 0 0 3px rgba(204,0,0,.1);
  }

  /* ── Submit button ── */
  .btn {
    width: 100%;
    padding: 12px;
    background: var(--red);
    color: #fff;
    border: none;
    border-radius: 6px;
    font-family: 'Inter', sans-serif;
    font-size: 14px;
    font-weight: 600;
    cursor: pointer;
    transition: background .15s;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    margin-top: 4px;
  }

  .btn:hover  { background: var(--red-dk); }
  .btn:active { transform: scale(.99); }
  .btn:disabled {
    background: #BDBDBD;
    cursor: not-allowed;
  }

  /* ── Status ── */
  #status {
    margin-top: 14px;
    border-radius: 6px;
    padding: 10px 14px;
    font-size: 13px;
    display: none;
    align-items: center;
    gap: 8px;
  }

  #status.loading {
    display: flex;
    background: #FFF8E1;
    border: 1px solid #FFE082;
    color: #795548;
  }

  #status.success {
    display: flex;
    background: #E8F5E9;
    border: 1px solid #A5D6A7;
    color: var(--success);
  }

  #status.error {
    display: flex;
    background: #FFEBEE;
    border: 1px solid #EF9A9A;
    color: #C62828;
  }

  .spinner {
    width: 14px; height: 14px;
    border: 2px solid currentColor;
    border-top-color: transparent;
    border-radius: 50%;
    animation: spin .7s linear infinite;
    flex-shrink: 0;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  /* ── Info row ── */
  .info-row {
    display: flex;
    gap: 10px;
    margin-top: 16px;
    flex-wrap: wrap;
    justify-content: center;
  }

  .info-chip {
    display: flex;
    align-items: center;
    gap: 5px;
    font-size: 11px;
    color: var(--muted);
  }

  .info-chip::before {
    content: '';
    display: inline-block;
    width: 14px;
    height: 14px;
    background-size: contain;
    background-repeat: no-repeat;
    opacity: .5;
  }

  /* ── Responsive ── */
  @media (max-width: 480px) {
    .fields { grid-template-columns: 1fr; }
    .card   { padding: 20px 18px; }
  }
</style>
</head>
<body>
<div class="wrapper">

  <header class="header">
    <span class="header-badge">Sistema de Relatórios</span>
    <h1>DER-MG &mdash; Educação para o Trânsito</h1>
    <p>Gere relatórios em PDF ou Excel com dados atualizados em tempo real.</p>
  </header>

  <div class="card">
    <form id="form">

      <div class="fields">
        <div class="field">
          <label for="ano">Ano</label>
          <select id="ano" name="ano">
            <option value="2023">2023</option>
            <option value="2024">2024</option>
            <option value="2025">2025</option>
            <option value="2026" selected>2026</option>
            <option value="2027">2027</option>
            <option value="2028">2028</option>
            <option value="2029">2029</option>
            <option value="2030">2030</option>
            <option value="2031">2031</option>
            <option value="2032">2032</option>
            <option value="2033">2033</option>
            <option value="2034">2034</option>
            <option value="2035">2035</option>
          </select>
        </div>

        <div class="field">
          <label for="mes">Mês</label>
          <select id="mes" name="mes">
            <option value="1">Janeiro</option>
            <option value="2">Fevereiro</option>
            <option value="3">Março</option>
            <option value="4">Abril</option>
            <option value="5">Maio</option>
            <option value="6">Junho</option>
            <option value="7">Julho</option>
            <option value="8">Agosto</option>
            <option value="9">Setembro</option>
            <option value="10">Outubro</option>
            <option value="11">Novembro</option>
            <option value="12">Dezembro</option>
          </select>
        </div>

        <div class="field full">
          <label for="tipo_relatorio">Tipo de Relatório</label>
          <select id="tipo_relatorio" name="tipo_relatorio">
            <option value="mensal">Mensal — ações do mês selecionado</option>
            <option value="acumulado">Acumulado — de janeiro até o mês selecionado</option>
            <option value="anual">Anual — visão completa do ano</option>
            <option value="metas">Metas — acompanhamento da meta anual</option>
            <option value="regional">Regional — um mês por página, todos os meses</option>
            <option value="sem_regional">Sem Regional — consolidado sem divisão</option>
            <option value="excel">Excel — planilha com todas as abas</option>
          </select>
        </div>
      </div>

      <button type="submit" class="btn" id="btn">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
          <polyline points="7 10 12 15 17 10"/>
          <line x1="12" y1="15" x2="12" y2="3"/>
        </svg>
        Gerar e Baixar Relatório
      </button>

      <div id="status" role="status" aria-live="polite"></div>

    </form>
  </div>



</div><!-- /wrapper -->

<script>
(function() {
  const mesAtual = new Date().getMonth() + 1;
  const select   = document.getElementById('mes');
  for (const opt of select.options) {
    if (parseInt(opt.value) === mesAtual) { opt.selected = true; break; }
  }
})();

const form   = document.getElementById('form');
const btn    = document.getElementById('btn');
const status = document.getElementById('status');

function setStatus(type, icon, msg) {
  status.className = type;
  status.innerHTML = icon + '<span>' + msg + '</span>';
}

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  const ano  = document.getElementById('ano').value;
  const mes  = document.getElementById('mes').value;
  const tipo = document.getElementById('tipo_relatorio').value;

  btn.disabled = true;
  setStatus('loading', '<div class="spinner"></div>', 'Buscando dados e gerando relatório, aguarde…');

  try {
    const body = new URLSearchParams({ ano, mes, tipo_relatorio: tipo });
    const resp = await fetch('/gerar', { method: 'POST', body });

    if (!resp.ok) {
      let msg = `Erro ${resp.status}`;
      try { const err = await resp.json(); msg = err.detail || msg; } catch (_) {}
      throw new Error(msg);
    }

    const blob     = await resp.blob();
    const url      = URL.createObjectURL(blob);
    const filename = decodeURIComponent(
      resp.headers.get('Content-Disposition')?.split("filename*=UTF-8''")[1]
      || resp.headers.get('Content-Disposition')?.split('filename=')[1]?.replace(/"/g, '')
      || 'relatorio'
    );

    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);

    setStatus('success',
      '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>',
      'Download iniciado com sucesso! (' + filename + ')'
    );
  } catch (err) {
    setStatus('error',
      '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>',
      'Erro: ' + err.message
    );
  } finally {
    btn.disabled = false;
  }
});
</script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
async def index():
    return HTML_PAGE


@app.post("/gerar")
async def gerar(
    ano:             int  = Form(...),
    mes:             int  = Form(..., ge=1, le=12),
    tipo_relatorio:  str  = Form(...),
):
    tipos_validos = {"mensal", "acumulado", "anual", "metas", "regional", "sem_regional", "excel"}
    if tipo_relatorio not in tipos_validos:
        return JSONResponse(status_code=400, content={"detail": f"Tipo '{tipo_relatorio}' inválido."})

    try:
        df    = carregar_dados()
        dados = DadosAgregados(df)

        df_verificacao = (
            dados.get_mensal(ano, mes)
            if tipo_relatorio in {"mensal", "acumulado", "excel"}
            else dados.get_anual(ano)
        )
        if df_verificacao.empty:
            nome_mes = MESES_PT.get(mes, str(mes))
            return JSONResponse(status_code=404, content={"detail": f"Nenhum dado para {nome_mes}/{ano}."})

        caminho = _gerar_relatorio(tipo_relatorio, dados, ano, mes, str(OUTPUT_DIR))

        media_type = (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if tipo_relatorio == "excel" else "application/pdf"
        )
        nome_arquivo = Path(caminho).name

        return FileResponse(
            path=caminho,
            media_type=media_type,
            filename=nome_arquivo,
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{nome_arquivo}"},
        )

    except RuntimeError as e:
        return JSONResponse(status_code=502, content={"detail": f"Erro ao acessar API: {str(e)}"})
    except Exception:
        traceback.print_exc()
        return JSONResponse(status_code=500, content={"detail": "Erro interno ao gerar o relatório."})


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
