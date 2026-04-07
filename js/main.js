// ---- SUPABASE CONFIG ----
      const SB_URL = "https://mhrnrhtdgdpdspmoayem.supabase.co".trim();
      const SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1ocm5yaHRkZ2RwZHNwbW9heWVtIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzUwNDMxMTMsImV4cCI6MjA5MDYxOTExM30.tGwVG81PZZdSgyZeLPI1q3Kcrm39pLV7REPgzkqIHYQ".trim();
      const _supabase = supabase.createClient(SB_URL, SB_KEY);

      const WHITELIST = [
        'denise.maliniak@globalsrv.com.br',
        'caroline.gomes@globalsrv.com.br',
        'marcelo.pereira@globalsrv.com.br'
      ];

      let actions = [];
      let obrasList = [];
      let isEditor = false;
      let userEmail = null;
      let currentObra = null;
      let editingIdx = null;
      let editingObraNome = null;
      let planoPage = 0;
      const PER_PAGE = 20;
      const fmt = v => v >= 1000000 ? 'R$ ' + (v / 1000000).toFixed(1) + 'M' : 'R$ ' + (v / 1000).toFixed(0) + 'k';



      // ---- AUTH LOGIC ----
      async function signInWithMicrosoft() {
        const { data, error } = await _supabase.auth.signInWithOAuth({
          provider: 'azure',
          options: {
            scopes: 'email profile',
            redirectTo: window.location.origin + window.location.pathname
          }
        });
        if (error) {
          // Fallback to 'microsoft' provider if 'azure' fails (sometimes configured differently)
          await _supabase.auth.signInWithOAuth({
            provider: 'microsoft',
            options: { scopes: 'email profile' }
          });
        }
      }

      async function signOut() {
        await _supabase.auth.signOut();
        window.location.reload();
      }

      async function checkUserSession() {
        try {
          const { data: { session } } = await _supabase.auth.getSession();

          const loginScreen = document.getElementById('login-screen');
          const appContainer = document.getElementById('app-container');

          if (session && session.user) {
            const user = session.user;
            userEmail = user.email;

            // Get metadata from Microsoft
            const fullName = user.user_metadata?.full_name || user.email;
            const avatarUrl = user.user_metadata?.avatar_url || `https://ui-avatars.com/api/?name=${encodeURIComponent(fullName)}&background=random`;

            // Update UI
            document.getElementById('user-name').textContent = fullName;
            document.getElementById('user-avatar').src = avatarUrl;
            document.getElementById('user-avatar').alt = fullName;
            document.getElementById('user-profile-data').style.display = 'flex';

            // Hide login, show app
            loginScreen.style.display = 'none';
            appContainer.style.display = 'block';

            // Check role in Supabase
            const { data: prof, error: profErr } = await _supabase
              .from('profissionais')
              .select('funcao')
              .ilike('email', userEmail.trim())
              .maybeSingle();

            if (!profErr && prof) {
              isEditor = (prof.funcao || '').toLowerCase().trim() === 'editor';
              console.log("Perfil carregado:", prof.funcao, "isEditor:", isEditor);
            } else {
              isEditor = false;
              if (profErr) console.warn("Erro ao buscar profissional:", profErr);
            }

            document.getElementById('user-role').textContent = isEditor ? 'Editor' : 'Visualizador';
            updateEditorVisibility();

            // Initial data fetch
            fetchFromSupabase();
          } else {
            // No session
            loginScreen.style.display = 'flex';
            appContainer.style.display = 'none';
            isEditor = false;
            updateEditorVisibility();
          }
        } catch (err) {
          console.error("Erro na sessão:", err);
        }
      }

      function updateEditorVisibility() {
        const display = isEditor ? '' : 'none';
        document.querySelectorAll('.editor-only').forEach(el => {
          el.style.display = display;
        });
      }

      // ---- UTILS ----
      function getTodayStr() {
        const d = new Date();
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return `${y}-${m}-${dd}`;
      }
      function normStatus(s) {
        return (s || '').toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      }
      function fmtMoney(v) {
        return 'R$ ' + (v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      }
      function fmtDate(d) {
        if (!d) return '—';
        const [y, m, dd] = d.split('-');
        return `${dd}/${m}/${y}`;
      }
      function statusBadge(s, dataPrev) {
        const todayStr = getTodayStr();
        const sn = normStatus(s);
        const isLate = sn !== 'concluido' && dataPrev && dataPrev < todayStr;

        if (isLate) return `<span class="badge badge-red">Atrasado</span>`;
        if (sn === 'concluido') return `<span class="badge badge-c">Concluído</span>`;
        if (sn === 'em andamento') return `<span class="badge badge-a">Em Andamento</span>`;
        if (sn === 'pendente') return `<span class="badge badge-purple">Pendente</span>`;
        if (sn === 'congelado') return `<span class="badge badge-g">Congelado</span>`;
        if (sn === 'atrasado') return `<span class="badge badge-red">Atrasado</span>`;
        return `<span class="badge">${s || ''}</span>`;
      }
      function progColor(pct) {
        if (pct >= 0.8) return 'green';
        if (pct >= 0.3) return 'yellow';
        return 'red';
      }
      // ---- CRITÉRIOS DE STATUS E ESTILO ----
      // Determina as cores e labels baseados no status e datas
      function situacaoBadge(s) {
        if (!s) return '';
        const sl = s.toLowerCase();
        if (sl.includes('crítico')) return `<span class="badge badge-crit">Crítico</span>`;
        if (sl.includes('progresso')) return `<span class="badge badge-prog">Em Progresso</span>`;
        if (sl.includes('concluído')) return `<span class="badge badge-c">Concluído</span>`;
        return `<span class="badge">${s}</span>`;
      }
      function getObrasStats() {
        const obras = {};
        actions.forEach(a => {
          if (!obras[a.obra]) obras[a.obra] = { obra: a.obra, gestor: a.gestor, prioridade: a.prioridade, retencao: a.valor_retido, conc: 0, and: 0, pend: 0, cong: 0, ag: 0, total: 0 };
          const o = obras[a.obra];
          const sn = normStatus(a.status);
          if (sn === 'concluido') o.conc++;
          else if (sn === 'em andamento') o.and++;
          else if (sn === 'congelado') o.cong++;
          else if (sn === 'aguardando') o.ag++;
          else o.pend++;
          o.total++;
        });
        return Object.values(obras);
      }

      // ---- PAGES ----
      function goPage(p) {
        document.querySelectorAll('.page').forEach(el => el.classList.remove('active'));
        document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
        document.getElementById('page-' + p).classList.add('active');
        const btn = document.querySelector(`.nav-item[onclick="goPage('${p}')"]`);
        if (btn) btn.classList.add('active');
        if (p === 'dashboard') { renderDash(); setTimeout(renderCharts, 50); }
        if (p === 'plano') renderPlano();
        if (p === 'financeiro') renderFin();
      }
      function goObra(obra) {
        currentObra = obra;
        document.querySelectorAll('.page').forEach(el => el.classList.remove('active'));
        document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
        document.getElementById('page-obra').classList.add('active');
        renderObra(obra);
      }

      // ---- RENDER DASH ----
      function renderDash() {
        const obras = getObrasStats();
        const todayStr = getTodayStr();
        const tbody = document.getElementById('dash-tbody');
        tbody.innerHTML = obras.map(o => {
          const pct = o.total ? o.conc / o.total : 0;
          const color = progColor(pct);
          // Status real: 100% = Concluído, 0 sem atraso = Crítico, qualquer atraso tb crítico
          const sit = pct === 1 ? '✅ Concluído' : pct > 0 ? '⚠️ Em Progresso' : '🔴 Crítico';
          const pb = `<div style="display:flex;align-items:center;gap:8px"><div class="prog-wrap" style="width:80px"><div class="prog-fill ${color}" style="width:${Math.round(pct * 100)}%"></div></div><span class="prog-label">${Math.round(pct * 100)}%</span></div>`;
          const pNum = (o.prioridade || '').match(/\d+/)?.[0] || '1';
          const pri = `<span class="tag-p${pNum}">P${pNum}</span>`;
          return `<tr>
      <td class="td-obra">${o.obra}</td>
      <td style="color:var(--text2)">${o.gestor}</td>
      <td>${pri}</td>
      <td style="text-align:center;color:var(--green);font-family:'JetBrains Mono',monospace">${o.conc}</td>
      <td style="text-align:center;color:var(--yellow);font-family:'JetBrains Mono',monospace">${o.and}</td>
      <td style="text-align:center;color:var(--gray);font-family:'JetBrains Mono',monospace">${o.pend}</td>
      <td style="text-align:center;color:var(--blue);font-family:'JetBrains Mono',monospace">${o.cong}</td>
      <td style="text-align:center;font-family:'JetBrains Mono',monospace;font-weight:600">${o.total}</td>
      <td>${pb}</td>
      <td class="td-money">${fmtMoney(o.retencao)}</td>
      <td>${situacaoBadge(sit)}</td>
      <td><button class="btn btn-ghost" style="padding:5px 12px;font-size:12px" onclick="goObra('${o.obra.replace(/'/g, "\\'")}')">Ver →</button></td>
    </tr>`;
        }).join('');
      }

      // ---- RENDER PLANO ----
      let planoFiltered = [];
      function renderPlano() {
        // populate selects once
        const obras = [...new Set(actions.map(a => a.obra))];
        const resps = [...new Set(actions.map(a => a.responsavel).filter(Boolean))].sort();
        const fObra = document.getElementById('f-obra');
        if (fObra.options.length <= 1) {
          obras.forEach(o => { const op = document.createElement('option'); op.value = o; op.textContent = o; fObra.appendChild(op); });
        }
        const fResp = document.getElementById('f-resp');
        if (fResp.options.length <= 1) {
          resps.forEach(r => { const op = document.createElement('option'); op.value = r; op.textContent = r; fResp.appendChild(op); });
        }
        filterPlano();
      }
      // Lógica de filtragem da página Plano de Ação
      function filterPlano() {
        const search = document.querySelector('#page-plano .search-inp').value.toLowerCase();
        const obra = document.getElementById('f-obra').value;
        const status = document.getElementById('f-status').value;
        const resp = document.getElementById('f-resp').value;
        const todayStr = getTodayStr();

        planoFiltered = actions.filter(a => {
          // Filtro por obra
          if (obra && a.obra !== obra) return false;

          // LOGICA DE STATUS: Aqui garantimos que o filtro combine com os cards do Dashboard
          if (status) {
            const sn = normStatus(a.status);
            const tn = normStatus(status);
            // Definição de ATRASADO: Não concluído e data vencida
            const isLate = (normStatus(a.status) !== 'concluido' && a.data_prevista && a.data_prevista < todayStr);

            if (tn === 'atrasado') {
              if (!isLate) return false;
            } else if (tn === 'em andamento') {
              // Só mostra "Em Andamento" se NÃO estiver atrasado
              if (sn !== 'em andamento' || isLate) return false;
            } else if (tn === 'congelado') {
              // Só mostra "Congelado" se NÃO estiver atrasado
              if (sn !== 'congelado' || isLate) return false;
            } else if (tn === 'pendente') {
              // Só mostra "Pendente" se NÃO estiver atrasado
              if (sn !== 'pendente' || isLate) return false;
            } else {
              if (sn !== tn) return false;
            }
          }

          // Filtro por responsável e busca textual
          if (resp && a.responsavel !== resp) return false;
          if (search && !a.acao.toLowerCase().includes(search) && !(a.responsavel || '').toLowerCase().includes(search)) return false;
          
          return true;
        });
        planoPage = 0;
        renderPlanoPage();
      }
      function renderPlanoPage() {
        const start = planoPage * PER_PAGE;
        const slice = planoFiltered.slice(start, start + PER_PAGE);
        const edCols = isEditor ? '<th>Ações</th>' : '';
        const tbody = document.getElementById('plano-tbody');
        tbody.innerHTML = slice.length ? slice.map(a => {
          const idx = actions.indexOf(a);
          const editBtn = isEditor ? `<td style="white-space:nowrap"><button class="btn btn-ghost" style="padding:4px 10px;font-size:11px" onclick="openEditModal(${idx})">Editar</button></td>` : '';
          return `<tr>
      <td class="td-mono">${a.num}</td>
      <td class="td-obra" style="max-width:130px">${a.obra}</td>
      <td><span class="tipo-badge">${a.tipo || ''}</span></td>
      <td class="td-acao">${a.acao}</td>
      <td style="color:var(--text2)">${a.responsavel || '—'}</td>
      <td style="color:var(--text3)">${a.apoio || '—'}</td>
      <td>${statusBadge(a.status, a.data_prevista)}</td>
      <td class="td-mono">${fmtDate(a.data_prevista)}</td>
      <td class="td-mono">${fmtDate(a.data_real)}</td>
      ${editBtn}
    </tr>`;
        }).join('') : '<tr><td colspan="10" class="empty">Nenhuma ação encontrada</td></tr>';

        // pagination
        const total = planoFiltered.length;
        const pages = Math.ceil(total / PER_PAGE);
        const pag = document.getElementById('plano-pag');
        let html = '';
        for (let i = 0; i < pages; i++) {
          html += `<button class="pag-btn ${i === planoPage ? 'active' : ''}" onclick="setPlanoPage(${i})">${i + 1}</button>`;
        }
        html += `<span class="pag-info">${total} ações encontradas</span>`;
        pag.innerHTML = html;
      }
      function setPlanoPage(p) { planoPage = p; renderPlanoPage(); }

      // ---- RENDER OBRA ----
      function renderObra(obra) {
        const oActions = actions.filter(a => a.obra === obra);
        const conc = oActions.filter(a => a.status === 'Concluído').length;
        const and = oActions.filter(a => a.status === 'Em Andamento').length;
        const pend = oActions.filter(a => a.status === 'Pendente' || a.status === 'PENDENTE').length;
        const cong = oActions.filter(a => (a.status || '').toUpperCase() === 'CONGELADO').length;
        const pct = oActions.length ? conc / oActions.length : 0;
        const retencao = oActions[0]?.valor_retido || 0;
        const gestor = oActions[0]?.gestor || '—';
        const pri = oActions[0]?.prioridade || '—';

        document.getElementById('obra-header-content').innerHTML = `
    <div class="page-header"><div class="page-title">${obra}</div></div>
    <div class="obra-header">
      <div class="obra-kpi"><div class="okpi-lbl">Gestor</div><div class="okpi-val">${gestor}</div></div>
      <div class="obra-kpi"><div class="okpi-lbl">Retenção</div><div class="okpi-val money">${fmtMoney(retencao)}</div></div>
      <div class="obra-kpi"><div class="okpi-lbl">Prioridade</div><div class="okpi-val">${pri}</div></div>
      <div class="obra-kpi"><div class="okpi-lbl">Progresso</div><div class="okpi-val">${Math.round(pct * 100)}% <span style="font-size:12px;color:var(--text3)">(${conc}/${oActions.length})</span></div></div>
      <div class="obra-kpi">
        <div class="okpi-lbl">Status por tipo</div>
        <div style="display:flex;gap:10px;margin-top:4px;flex-wrap:wrap">
          <span style="font-size:12px;color:var(--green)">✓ ${conc}</span>
          <span style="font-size:12px;color:var(--yellow)">⟳ ${and}</span>
          <span style="font-size:12px;color:var(--gray)">○ ${pend}</span>
          ${cong ? `<span style="font-size:12px;color:var(--blue)">❄ ${cong}</span>` : ''}
        </div>
      </div>
    </div>`;
        filterObra();
      }
      function filterObra() {
        if (!currentObra) return;
        const search = document.querySelector('#page-obra .search-inp').value.toLowerCase();
        const status = document.getElementById('f-obra-status').value;
        const filtered = actions.filter(a => {
          if (a.obra !== currentObra) return false;
          if (status && (a.status || '').toLowerCase() !== status.toLowerCase()) return false;
          if (search && !a.acao.toLowerCase().includes(search)) return false;
          return true;
        });
        const tbody = document.getElementById('obra-tbody');
        tbody.innerHTML = filtered.length ? filtered.map(a => {
          const idx = actions.indexOf(a);
          const editBtns = isEditor ? `<td style="white-space:nowrap">
      <button class="btn btn-ghost" style="padding:4px 9px;font-size:11px" onclick="openEditModal(${idx})">Editar</button>
    </td>` : '';
          return `<tr>
      <td class="td-mono">${a.num}</td>
      <td><span class="tipo-badge">${a.tipo || ''}</span></td>
      <td class="td-acao">${a.acao}</td>
      <td style="color:var(--text2)">${a.responsavel || '—'}</td>
      <td style="color:var(--text3)">${a.apoio || '—'}</td>
      <td>${statusBadge(a.status)}</td>
      <td class="td-mono">${fmtDate(a.data_prevista)}</td>
      <td class="td-mono">${fmtDate(a.data_real)}</td>
      <td style="font-size:12px;color:var(--text3)">—</td>
      ${editBtns}
    </tr>`;
        }).join('') : '<tr><td colspan="10" class="empty">Nenhuma ação encontrada</td></tr>';
      }

      // ---- RENDER FIN ----
      function renderFin() {
        const obras = getObrasStats();
        let totalRet = 0;
        const tbody = document.getElementById('fin-tbody');
        tbody.innerHTML = obras.map(o => {
          totalRet += o.retencao;
          const pct = o.total ? o.conc / o.total : 0;
          const sit = pct === 1 ? '✅ Concluído' : pct > 0 ? '⚠️ Em Progresso' : '🔴 Crítico';
          const pNum = (o.prioridade || '').match(/\d+/)?.[0] || '1';
          const pri = `<span class="tag-p${pNum}">P${pNum}</span>`;
          const pb = `<div style="display:flex;align-items:center;gap:8px"><div class="prog-wrap" style="width:60px"><div class="prog-fill ${progColor(pct)}" style="width:${Math.round(pct * 100)}%"></div></div><span class="td-mono">${Math.round(pct * 100)}%</span></div>`;
          return `<tr>
      <td class="td-obra">${o.obra}</td>
      <td style="color:var(--text2)">${o.gestor}</td>
      <td>${pri}</td>
      <td class="td-money">${fmtMoney(o.retencao)}</td>
      <td style="text-align:center;font-family:'DM Mono',monospace">${o.total}</td>
      <td>${pb}</td>
      <td>${situacaoBadge(sit)}</td>
    </tr>`;
        }).join('');
        tbody.innerHTML += `<tr class="fin-total">
    <td colspan="3">TOTAL GERAL</td>
    <td class="td-money" style="font-size:15px">${fmtMoney(totalRet)}</td>
    <td style="text-align:center;font-family:'DM Mono',monospace">${actions.length}</td>
    <td colspan="2"></td>
  </tr>`;
      }

      // ---- EDITOR MODE ----
      function toggleMode() {
        isEditor = !isEditor;
        const pill = document.getElementById('mode-toggle');
        const label = document.getElementById('mode-label');
        pill.classList.toggle('editor', isEditor);
        label.textContent = isEditor ? 'Modo Editor' : 'Modo Visualização';
        document.querySelectorAll('.editor-only').forEach(el => {
          el.style.display = isEditor ? '' : 'none';
        });
        // re-render current page
        const active = document.querySelector('.page.active');
        if (active?.id === 'page-plano') renderPlanoPage();
        if (active?.id === 'page-obra') filterObra();
      }

      // ---- MODAL ----
      function openEditModal(idx) {
        editingIdx = idx;
        const a = actions[idx];
        document.getElementById('modal-title').textContent = `Editar Ação #${a.num}`;
        document.getElementById('f-acao').value = a.acao || '';
        document.getElementById('f-tipo').value = a.tipo || 'Documentação';
        document.getElementById('f-modal-status').value = a.status || 'Pendente';
        document.getElementById('f-responsavel').value = a.responsavel || '';
        document.getElementById('f-email-responsavel').value = a.email_responsavel || '';
        document.getElementById('f-apoio').value = a.apoio || '';
        document.getElementById('f-data-prev').value = a.data_prevista || '';
        document.getElementById('f-data-real').value = a.data_real || '';
        document.getElementById('f-observacoes').value = a.observacoes || '';
        document.getElementById('modal-overlay').classList.add('open');
      }
      function openAddModal(obra) {
        editingIdx = null;
        document.getElementById('modal-title').textContent = 'Nova Ação' + (obra ? ` — ${obra}` : '');
        document.getElementById('f-acao').value = '';
        document.getElementById('f-tipo').value = 'Documentação';
        document.getElementById('f-modal-status').value = 'Pendente';
        document.getElementById('f-responsavel').value = '';
        document.getElementById('f-email-responsavel').value = '';
        document.getElementById('f-apoio').value = '';
        document.getElementById('f-data-prev').value = '';
        document.getElementById('f-data-real').value = '';
        document.getElementById('f-observacoes').value = '';
        document.getElementById('modal-overlay').classList.add('open');
      }
      function closeModal(e) {
        if (!e || e.target === document.getElementById('modal-overlay')) {
          document.getElementById('modal-overlay').classList.remove('open');
        }
      }
      async function saveAction() {
        if (!isEditor) return showToast("Permissão negada");

        const saveBtn = document.getElementById('save-btn');

        // ---- VALIDATION ----
        const reqFields = {
          'Ação': document.getElementById('f-acao').value,
          'Tipo': document.getElementById('f-tipo').value,
          'Status': document.getElementById('f-modal-status').value,
          'Responsável': document.getElementById('f-responsavel').value,
          'E-mail Resp.': document.getElementById('f-email-responsavel').value,
          'Data Prevista': document.getElementById('f-data-prev').value,
          'Obra': document.getElementById('f-obra-modal').value,
          'Gestor': document.getElementById('f-gestor').value
        };

        for (const [name, val] of Object.entries(reqFields)) {
          if (!val || val.trim() === '') {
            showToast(`Preencha o campo obrigatório: ${name} ⚠️`);
            return;
          }
        }

        saveBtn.disabled = true;
        saveBtn.textContent = "Sincronizando...";

        const previousStatus = editingIdx !== null ? actions[editingIdx].status : null;
        const currentStatus = document.getElementById('f-modal-status').value;

        const currentData = {
          num: editingIdx !== null ? Number(actions[editingIdx].num) : (actions.reduce((max, a) => Math.max(max, Number(a.num) || 0), 0) + 1),
          obra: document.getElementById('f-obra-modal').value,
          tipo: document.getElementById('f-tipo').value,
          status: currentStatus,
          responsavel: document.getElementById('f-responsavel').value,
          email_responsavel: document.getElementById('f-email-responsavel').value,
          apoio: document.getElementById('f-apoio').value,
          data_prevista: document.getElementById('f-data-prev').value || null,
          data_real: document.getElementById('f-data-real').value || null,
          acao: document.getElementById('f-acao').value,
          observacoes: document.getElementById('f-observacoes').value,
          gestor: document.getElementById('f-gestor').value,
          prioridade: document.getElementById('f-prioridade').value,
          valor_retido: Number(document.getElementById('f-valor-retido').value) || 0
        };

        if (editingIdx !== null) {
          currentData.id = actions[editingIdx].id;
        }

        console.log("Enviando para Supabase:", currentData);

        // 1. Instant feedback (Optimistic Update)
        closeModal();
        showToast("Atualizando plataforma...");

        const { data, error } = await _supabase.from('acoes').upsert(currentData).select();

        if (error) {
          console.error("Erro ao salvar:", error);
          showToast("Erro ao conectar com Supabase: " + error.message);
          saveBtn.disabled = false;
          saveBtn.textContent = "Salvar alterações";
          return;
        }

        // 2. Log history if status or observations changed
        const previousObs = editingIdx !== null ? (actions[editingIdx].observacoes || '') : '';
        const currentObs = currentData.observacoes || '';

        if (previousStatus !== currentStatus || previousObs !== currentObs) {
          const changelog = [];
          if (previousStatus !== currentStatus) changelog.push(`Status: ${previousStatus || 'Novo'} → ${currentStatus}`);
          if (previousObs !== currentObs) changelog.push(`Obs atualizada`);

          await _supabase.from('historico_acoes').insert({
            num: currentData.num,
            alterado_por: userEmail,
            status_anterior: previousStatus || 'Nova Ação',
            status_novo: currentStatus,
            acao_desc: changelog.join(' | '),
            observacao: currentObs
          });
        }

        showToast("Alterações feitas com sucesso ✅");
        await fetchFromSupabase();

        // Reset button state
        saveBtn.disabled = false;
        saveBtn.textContent = "Salvar alterações";
      }

      async function fetchFromSupabase() {
        const statusEl = document.getElementById('sheets-status');
        if (statusEl) statusEl.textContent = '⌛ Sincronizando...';

        const [acoesRes, obrasRes] = await Promise.all([
          _supabase.from('acoes').select('*').order('num', { ascending: true }),
          _supabase.from('obras').select('*').order('nome', { ascending: true })
        ]);

        if (acoesRes.error) {
          console.error("Erro no fetch:", acoesRes.error);
          if (statusEl) statusEl.textContent = '❌ Supabase Off';
          return;
        }

        actions = acoesRes.data;
        obrasList = obrasRes.data || [];
        if (statusEl) statusEl.textContent = '✅ Supabase On';

        const active = document.querySelector('.page.active');
        if (active?.id === 'dashboard' || active?.id === 'page-dashboard') { renderDash(); renderCharts(); }
        else if (active?.id === 'plano' || active?.id === 'page-plano') renderPlano();
        else if (active?.id === 'financeiro' || active?.id === 'page-financeiro') renderFin();
        else if (active?.id === 'obra' || active?.id === 'page-obra') renderObra(currentObra);
      }

      async function renderTimeline(num) {
        const timelineWrap = document.getElementById('timeline-section');
        const timelineList = document.getElementById('timeline-list');
        timelineWrap.style.display = 'block';
        timelineList.innerHTML = '<div style="font-size:12px;color:var(--text3)">Carregando histórico...</div>';

        const { data, error } = await _supabase
          .from('historico_acoes')
          .select('*')
          .eq('num', num)
          .order('data_alteracao', { ascending: false });

        if (error || !data.length) {
          timelineList.innerHTML = '<div style="font-size:12px;color:var(--text3)">Nenhum histórico encontrado.</div>';
          return;
        }

        timelineList.innerHTML = data.map(h => {
          const obsHtml = h.observacao ? `<div style="font-size:11px;color:var(--text2);margin-top:4px;font-style:italic">"${h.observacao}"</div>` : '';
          return `
          <div style="background:var(--bg);padding:10px;border-radius:8px;border:1px solid var(--border)">
            <div style="display:flex;justify-content:space-between;font-size:11px;color:var(--text3);margin-bottom:4px">
              <span>${h.alterado_por || 'Sistema'}</span>
              <span>${new Date(h.data_alteracao).toLocaleString('pt-BR')}</span>
            </div>
            <div style="font-weight:600;font-size:12px">${h.acao_desc || 'Alteração realizada'}</div>
            ${obsHtml}
          </div>`;
        }).join('');
      }

      // ---- MODAL ----
      function suggestObraInfo() {
        const obra = document.getElementById('f-obra-modal').value;
        if (!obra) return;
        const match = actions.find(a => a.obra === obra);
        if (match) {
          document.getElementById('f-gestor').value = match.gestor || '';
          document.getElementById('f-valor-retido').value = match.valor_retido || 0;
          document.getElementById('f-prioridade').value = match.prioridade || 'PRIORIDADE 1';
        }
      }

      async function openEditModal(idx) {
        editingIdx = idx;
        const a = actions[idx];

        // Populate Obra list in modal
        const obras = [...new Set(actions.map(act => act.obra))].sort();
        const fObra = document.getElementById('f-obra-modal');
        fObra.innerHTML = '<option value="">Selecione a Obra...</option>' + obras.map(o => `<option value="${o}">${o}</option>`).join('');
        fObra.value = a.obra;
        fObra.disabled = true; // Lock obra on edit

        document.getElementById('modal-title').textContent = `Editar Ação #${a.num}`;
        document.getElementById('f-acao').value = a.acao || '';
        document.getElementById('f-tipo').value = a.tipo || 'Documentação';
        document.getElementById('f-modal-status').value = a.status || 'Pendente';
        document.getElementById('f-responsavel').value = a.responsavel || '';
        document.getElementById('f-email-responsavel').value = a.email_responsavel || '';
        document.getElementById('f-gestor').value = a.gestor || '';
        document.getElementById('f-prioridade').value = a.prioridade || 'PRIORIDADE 1';
        document.getElementById('f-apoio').value = a.apoio || '';
        document.getElementById('f-valor-retido').value = a.valor_retido || 0;
        document.getElementById('f-data-prev').value = a.data_prevista || '';
        document.getElementById('f-data-real').value = a.data_real || '';
        document.getElementById('f-observacoes').value = a.observacoes || '';

        document.getElementById('modal-overlay').classList.add('open');
        await renderTimeline(a.num);

        if (isEditor) {
          const saveBtn = document.getElementById('save-btn');
          saveBtn.style.display = 'block';
          saveBtn.disabled = false;
          saveBtn.textContent = 'Salvar alterações';
        }
      }

      function openAddModal(obra) {
        editingIdx = null;

        // Populate Obra list in modal
        const obras = [...new Set(actions.map(act => act.obra))].sort();
        const fObra = document.getElementById('f-obra-modal');
        fObra.innerHTML = '<option value="">Selecione a Obra...</option>' + obras.map(o => `<option value="${o}">${o}</option>`).join('');

        document.getElementById('modal-title').textContent = 'Nova Ação' + (obra ? ` — ${obra}` : '');
        document.getElementById('f-acao').value = '';
        document.getElementById('f-tipo').value = 'Documentação';
        document.getElementById('f-modal-status').value = 'Pendente';
        document.getElementById('f-responsavel').value = '';
        document.getElementById('f-email-responsavel').value = '';
        document.getElementById('f-gestor').value = '';
        document.getElementById('f-valor-retido').value = 0;
        document.getElementById('f-prioridade').value = 'PRIORIDADE 1';
        document.getElementById('f-apoio').value = '';
        document.getElementById('f-data-prev').value = '';
        document.getElementById('f-data-real').value = '';
        document.getElementById('f-observacoes').value = '';
        document.getElementById('timeline-section').style.display = 'none';

        if (obra) {
          fObra.value = obra;
          fObra.disabled = true;
          suggestObraInfo(); // Pre-fill gestor/valor
        } else {
          fObra.disabled = false;
        }

        if (isEditor) {
          const saveBtn = document.getElementById('save-btn');
          saveBtn.style.display = 'block';
          saveBtn.disabled = false;
          saveBtn.textContent = 'Salvar alterações';
        }

        document.getElementById('modal-overlay').classList.add('open');
      }

      // Initial fetch & auth check
      window.addEventListener('DOMContentLoaded', async () => {
        // Set dynamic dates on all screens
        const tds = fmtDate(getTodayStr());
        const topbarDate = document.getElementById('topbar-date');
        if (topbarDate) topbarDate.textContent = tds;
        const heroDate = document.getElementById('hero-base-date');
        if (heroDate) heroDate.textContent = tds;
        const finDate = document.getElementById('fin-base-date');
        if (finDate) finDate.textContent = tds;

        await fetchFromSupabase();
        await checkUserSession();
      });

      function showToast(msg) {
        const t = document.getElementById('toast');
        if (!t) {
          console.log("Toast:", msg);
          return;
        }
        t.textContent = msg || 'Alteração salva com sucesso';
        t.classList.add('show');
        setTimeout(() => t.classList.remove('show'), 2500);
      }

      function filterByStatus(statusName) {
        goPage('plano');
        setTimeout(() => {
          const select = document.getElementById('f-status');
          if (select) {
            select.value = statusName;
            filterPlano();
          }
        }, 50);
      }

      // ---- DELETE AÇÃO ----
      async function deleteAction() {
        if (!isEditor) return showToast('Permissão negada');
        if (editingIdx === null) return;
        const a = actions[editingIdx];
        const confirmed = window.confirm(`Tem certeza que deseja excluir a ação #${a.num}?\n"${a.acao}"\n\nEsta operação não pode ser desfeita.`);
        if (!confirmed) return;
        const { error } = await _supabase.from('acoes').delete().eq('id', a.id);
        if (error) return showToast('Erro ao excluir: ' + error.message);
        showToast('✅ Ação excluída com sucesso');
        closeModal();
        await fetchFromSupabase();
      }

      // ---- GESTÃO DE OBRAS ----
      // (obrasList e editingObraNome já declarados no topo)

      function openGestaoObras() {
        renderObrasList();
        document.getElementById('obra-form-wrap').style.display = 'none';
        document.getElementById('obras-overlay').classList.add('open');
      }

      function closeGestaoObras(e) {
        if (!e || e.target === document.getElementById('obras-overlay')) {
          document.getElementById('obras-overlay').classList.remove('open');
        }
      }

      function renderObrasList() {
        const wrap = document.getElementById('obras-list-wrap');
        if (!obrasList.length) {
          wrap.innerHTML = '<div class="empty">Nenhuma obra cadastrada ainda.</div>';
          return;
        }
        wrap.innerHTML = `<div style="display:flex;flex-direction:column;gap:10px">${obrasList.map(o => {
          const pNum = (o.prioridade || '').match(/\d+/)?.[0] || '1';
          return `
          <div style="background:var(--bg);border:1px solid var(--border);border-radius:10px;padding:14px 18px;display:flex;align-items:center;gap:12px">
            <div style="flex:1">
              <div style="font-weight:600;font-size:13px">${o.nome || o.name || ''}</div>
              <div style="font-size:12px;color:var(--text2);margin-top:3px">${o.gestor || '—'} &nbsp;•&nbsp; <span class="tag-p${pNum}">P${pNum}</span> &nbsp;•&nbsp; ${fmtMoney(o.valor_retido || 0)}</div>
            </div>
            <button class="btn btn-ghost" style="font-size:11px;padding:5px 12px" onclick="openEditObraForm('${(o.nome || o.name || '').replace(/'/g, "\\'")}')">Editar</button>
          </div>`;
        }).join('')}</div>`;
      }

      function openNovaObraForm() {
        editingObraNome = null;
        document.getElementById('obra-form-title').textContent = 'Nova Obra';
        document.getElementById('go-nome').value = '';
        document.getElementById('go-nome').disabled = false;
        document.getElementById('go-gestor').value = '';
        document.getElementById('go-prioridade').value = 'PRIORIDADE 1';
        document.getElementById('go-retencao').value = '';
        document.getElementById('go-status').value = 'Em Andamento';
        document.getElementById('go-save-btn').textContent = 'Salvar Obra';
        document.getElementById('obra-form-wrap').style.display = 'block';
      }

      function openEditObraForm(nome) {
        editingObraNome = nome;
        const o = obrasList.find(ob => (ob.nome || ob.name) === nome);
        if (!o) return;
        document.getElementById('obra-form-title').textContent = 'Editar Obra';
        document.getElementById('go-nome').value = o.nome || o.name || '';
        document.getElementById('go-nome').disabled = true;
        document.getElementById('go-gestor').value = o.gestor || '';
        document.getElementById('go-prioridade').value = o.prioridade || 'PRIORIDADE 1';
        document.getElementById('go-retencao').value = o.valor_retido || '';
        document.getElementById('go-status').value = o.status || 'Em Andamento';
        document.getElementById('go-save-btn').textContent = 'Atualizar Obra';
        document.getElementById('obra-form-wrap').style.display = 'block';
      }

      function cancelObraForm() {
        document.getElementById('obra-form-wrap').style.display = 'none';
      }

      async function saveObraGestao() {
        if (!isEditor) return showToast('Permissão negada');
        const nome = document.getElementById('go-nome').value.trim();
        const gestor = document.getElementById('go-gestor').value.trim();
        const prioridade = document.getElementById('go-prioridade').value;
        const valor_retido = Number(document.getElementById('go-retencao').value) || 0;
        const status = document.getElementById('go-status').value;
        if (!nome || !gestor) return showToast('⚠️ Preencha Nome e Gestor.');
        const saveBtn = document.getElementById('go-save-btn');
        saveBtn.disabled = true;
        saveBtn.textContent = 'Salvando...';
        const payload = { nome, gestor, prioridade, valor_retido, status };
        const { error } = await _supabase.from('obras').upsert(payload, { onConflict: 'nome' });
        if (error) {
          showToast('Erro: ' + error.message);
        } else {
          showToast('✅ Obra salva com sucesso!');
          cancelObraForm();
          await fetchFromSupabase();
          renderObrasList();
        }
        saveBtn.disabled = false;
        saveBtn.textContent = editingObraNome ? 'Atualizar Obra' : 'Salvar Obra';
      }

      // ---- CHARTS ----
      let chartDonut, chartPrio, chartObras, chartRet;

      function renderCharts() {
        const obras = getObrasStats();
        const todayStr = getTodayStr();

        // Show current base date in dashboard
        const dateIndicator = document.getElementById('base-date-indicator');
        if (dateIndicator) dateIndicator.textContent = `Base: ${fmtDate(todayStr)}`;

        // Helper to normalize status checks
        // Função auxiliar para contar tarefas por status de forma EXCLUSIVA
        // Uma tarefa só pode pertencer a UM grupo (atrasado tem prioridade sobre os outros)
        const getStatusCount = (target) => {
          return actions.filter(a => {
            const sn = normStatus(a.status);
            const tn = normStatus(target);
            const isLate = (sn !== 'concluido' && a.data_prevista && a.data_prevista < todayStr);

            // 1. Se o objetivo é contar ATRASADOS
            if (tn === 'atrasado') return isLate;
            
            // 2. Se a tarefa está ATRASADA, ela não entra em outras contagens (exceto se já concluída)
            if (isLate && tn !== 'concluido') return false;

            // 3. Verificação normal de status para itens no prazo
            if (tn === 'concluido') return sn === 'concluido';
            if (tn === 'pendente') return sn === 'pendente';
            if (tn === 'em andamento') return sn === 'em andamento';
            if (tn === 'congelado') return sn === 'congelado';

            return sn === tn;
          }).length;
        };

        const conc = getStatusCount('Concluído');
        const and = getStatusCount('Em Andamento');
        const pend = getStatusCount('Pendente');
        const atras = getStatusCount('Atrasado');
        const cong = getStatusCount('Congelado');
        const total = actions.length;

        // ---- CÁLCULOS GLOBAIS ----
        const totalRetencaoGlobal = obras.reduce((s, o) => s + (o.retencao || 0), 0);
        const pctGlobal = total > 0 ? Math.round((conc / total) * 100) : 0;

        // Atualização Dinâmica do Hero
        document.getElementById('hero-val').textContent = fmtMoney(totalRetencaoGlobal);
        document.getElementById('hero-pct').textContent = pctGlobal + '%';
        document.getElementById('hero-summary-label').textContent = `${conc} de ${total} ações concluídas`;

        // Atualização Dinâmica dos Cards KPI
        document.getElementById('kpi-total').textContent = total;
        document.getElementById('kpi-conc').textContent = conc;
        document.getElementById('kpi-and').textContent = and;
        document.getElementById('kpi-pend').textContent = pend;
        document.getElementById('kpi-atras').textContent = atras;

        const getPctStr = (v) => total > 0 ? Math.round((v / total) * 100) + '%' : '0%';
        document.getElementById('kpi-conc-pct').textContent = getPctStr(conc);
        document.getElementById('kpi-and-pct').textContent = getPctStr(and);
        document.getElementById('kpi-pend-pct').textContent = getPctStr(pend);
        document.getElementById('kpi-atras-pct').textContent = getPctStr(atras);

        const labelRet = document.getElementById('ch-total-ret-label');
        if (labelRet) labelRet.textContent = `Valor retido total: ${fmtMoney(totalRetencaoGlobal)}`;

        // ---- DONUT STATUS (Chart.js) ----
        if (chartDonut) chartDonut.destroy();
        chartDonut = new Chart(document.getElementById('chart-donut'), {
          type: 'doughnut',
          data: {
            labels: ['Concluído', 'Em Andamento', 'Pendente', 'Atrasado', 'Congelado'],
            datasets: [{
              data: [conc, and, pend, atras, cong],
              backgroundColor: ['#16a34a', '#d97706', '#7c3aed', '#dc2626', '#1d4ed8'],
              borderColor: '#ffffff', borderWidth: 3, hoverOffset: 8
            }]
          },
          options: {
            cutout: '72%', animation: { duration: 800, easing: 'easeOutQuart' },
            plugins: {
              legend: { display: false },
              tooltip: {
                callbacks: {
                  label: ctx => {
                    const pct = total > 0 ? Math.round((ctx.raw / total) * 100) : 0;
                    return ` ${ctx.label}: ${ctx.raw} ações (${pct}%)`;
                  }
                }
              }
            }
          }
        });
        document.getElementById('donut-total').textContent = total;

        // Custom legend for donut
        const dColors = ['#16a34a', '#d97706', '#7c3aed', '#dc2626', '#1d4ed8'];
        const dLabels = ['Concluído', 'Em Andamento', 'Pendente', 'Atrasado', 'Congelado'];
        const dVals = [conc, and, pend, atras, cong];
        document.getElementById('donut-legend').innerHTML = dLabels.map((l, i) => {
          if (dVals[i] === 0 && l === 'Congelado') return ''; // hide if zero and less important
          return `
    <div class="dl-item">
      <div class="dl-dot" style="background:${dColors[i]}"></div>
      <div class="dl-label">${l}</div>
      <div class="dl-val">${dVals[i]}<span class="dl-pct">${total > 0 ? Math.round(dVals[i] / total * 100) : 0}%</span></div>
    </div>`;
        }).join('');

        // ---- DONUT PRIORIDADE (Chart.js) ----
        const pData = [0, 0, 0, 0, 0, 0]; // P0-P5
        obras.forEach(o => {
          const n = parseInt((o.prioridade || '').match(/\d+/)?.[0]);
          if (!isNaN(n) && n >= 0 && n <= 5) pData[n] += o.retencao;
          else pData[1] += o.retencao; // default P1
        });

        if (chartPrio) chartPrio.destroy();
        chartPrio = new Chart(document.getElementById('chart-prio'), {
          type: 'doughnut',
          data: {
            labels: ['P0', 'P1', 'P2', 'P3', 'P4', 'P5'],
            datasets: [{
              data: pData,
              backgroundColor: ['#dc2626', '#1d4ed8', '#64748b', '#94a3b8', '#cbd5e1', '#e2e8f0'],
              borderColor: '#ffffff', borderWidth: 2, hoverOffset: 8
            }]
          },
          options: {
            cutout: '72%', animation: { duration: 800, easing: 'easeOutQuart' },
            plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ` ${ctx.label}: R$ ${ctx.raw.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` } } }
          }
        });

        const pTotal = pData.reduce((a, b) => a + b, 0);
        const pLabels = ['P0', 'P1', 'P2', 'P3', 'P4', 'P5'];
        const pColors = ['#dc2626', '#1d4ed8', '#64748b', '#94a3b8', '#cbd5e1', '#e2e8f0'];

        document.getElementById('prio-legend').innerHTML = pData.map((v, i) => {
          if (v === 0) return '';
          return `
    <div class="dl-item">
      <div class="dl-dot" style="background:${pColors[i]}"></div>
      <div class="dl-label">${pLabels[i]}</div>
      <div style="text-align:right">
        <div class="dl-val" style="font-size:11px">${fmt(v)}</div>
        <div class="dl-pct">${pTotal > 0 ? Math.round(v / pTotal * 100) : 0}%</div>
      </div>
    </div>`;
        }).join('');

        // ---- HORIZONTAL BARS — PROGRESSO ----
        const maxPct = 100;
        const hbarHTML = obras.map(o => {
          const pct = o.total ? Math.round(o.conc / o.total * 100) : 0;
          const color = pct === 100 ? '#16a34a' : pct > 0 ? '#d97706' : '#dc2626';
          const shortName = o.obra.replace('MOTIVA / ', '').replace(' (EPA)', '').replace('CCR ', '');
          return `<div class="hbar-item">
      <div class="hbar-header">
        <div class="hbar-label" title="${o.obra}">${shortName}</div>
        <div class="hbar-vals">
          <span class="hbar-pct" style="color:${color}">${pct}%</span>
          <span class="hbar-count">${o.conc}/${o.total} concluídas</span>
        </div>
      </div>
      <div class="hbar-track">
        <div class="hbar-fill" style="width:${pct}%;background:${color}"></div>
      </div>
    </div>`;
        }).join('');
        document.getElementById('hbar-obras').innerHTML = hbarHTML;

        // ---- HORIZONTAL BARS — RETENÇÃO ----
        const maxRet = Math.max(...obras.map(o => o.retencao));
        const retHTML = obras.map(o => {
          const w = maxRet > 0 ? Math.round(o.retencao / maxRet * 100) : 0;
          const money = 'R$ ' + o.retencao.toLocaleString('pt-BR', { minimumFractionDigits: 2 });
          const shortName = o.obra.replace('MOTIVA / ', '').replace(' (EPA)', '').replace('CCR ', '');
          return `<div class="retbar-item">
      <div class="retbar-header">
        <div class="retbar-name" title="${o.obra}">${shortName}</div>
        <div class="retbar-money">${money}</div>
      </div>
      <div class="retbar-track">
        <div class="retbar-fill" style="width:${w}%"></div>
      </div>
    </div>`;
        }).join('');
        document.getElementById('retbar-obras').innerHTML = retHTML;

        // unused chart vars — clear them
        if (chartObras) { try { chartObras.destroy(); } catch (e) { } chartObras = null; }
        if (chartRet) { try { chartRet.destroy(); } catch (e) { } chartRet = null; }
      }

      // Initial fetch & auth check
      window.addEventListener('DOMContentLoaded', async () => {
        // Initialize dynamic dates
        const tds = fmtDate(getTodayStr());
        document.getElementById('topbar-date').textContent = tds;
        document.getElementById('hero-base-date').textContent = tds;
        document.getElementById('fin-base-date').textContent = tds;

        await fetchFromSupabase();
        await checkUserSession();
      });
      _supabase.auth.onAuthStateChange((event, session) => {
        console.log("Auth event:", event);
        checkUserSession();
      });