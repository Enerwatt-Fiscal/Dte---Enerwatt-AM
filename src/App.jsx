import { useState, useEffect, useCallback } from "react";
import { supabase } from "./supabase";

// Exportação Excel via SheetJS (carregado via CDN no index.html)
function exportarExcel(filtradas, empresas, periodo) {
  try {
    const XLSX = window.XLSX;
    if (!XLSX) { alert("Biblioteca Excel não carregada. Recarregue a página."); return; }
    const nomeEmp = (id) => { const e = empresas.find(x => x.id === id); return e ? e.nome : id; };
    const num = (v) => Number(v||0);
    const wb = XLSX.utils.book_new();

    // ---- ABA 1: DADOS COMPLETOS ----
    const dadosCompletos = filtradas.map(n => ({
      "Empresa": nomeEmp(n.empresa),
      "Fornecedor": n.fornecedor || "",
      "Razão Social": n.razaoSocial || "",
      "Nº Nota": n.numNota || "",
      "CFOP": n.cfop || "",
      "Chave NF-e": n.chave || "",
      "Dt. Emissão": n.dtEmissao || "",
      "Dt. Apresentação": n.dtApresentacao || "",
      "Dt. Importação": n.dtImportacao || "",
      "Valor NF (R$)": num(n.valor),
      "Qtde Dias": num(n.qtdeDias),
      "Faixa de Prazo": num(n.qtdeDias) >= 60 ? "Crítico (60+d)" : num(n.qtdeDias) >= 25 ? "Atenção (25-59d)" : "OK (0-24d)",
      "Status": n.status || "",
      "Finalidade": n.finalidade || "",
      "Centro de Custo": n.centroCusto || "",
      "Responsável": n.responsavel || "",
      "No RM": n.noRM === true ? "Sim" : n.noRM === false ? "Não" : "-",
      "Taxa Reanálise (R$)": num(n.taxaReanalise),
      "Taxa Desembaraço (R$)": num(n.taxaDesembaraco),
      "ICMS Antecipado (R$)": num(n.icmsAntecipado),
      "Multa (R$)": num(n.multa),
      "Juros (R$)": num(n.juros),
      "Total Custos (R$)": num(n.taxaReanalise)+num(n.taxaDesembaraco)+num(n.icmsAntecipado)+num(n.multa10pct||0)+num(n.multa)+num(n.juros),
      "Observações": n.obs || "",
    }));
    const ws1 = XLSX.utils.json_to_sheet(dadosCompletos);
    ws1["!cols"] = [{wch:22},{wch:18},{wch:30},{wch:12},{wch:8},{wch:46},{wch:13},{wch:16},{wch:14},{wch:14},{wch:10},{wch:16},{wch:22},{wch:16},{wch:18},{wch:16},{wch:20},{wch:8},{wch:18},{wch:20},{wch:20},{wch:14},{wch:12},{wch:16},{wch:30}];
    XLSX.utils.book_append_sheet(wb, ws1, "Dados Completos");

    // ---- ABA 2: RESUMO POR EMPRESA ----
    const empMap = {};
    filtradas.forEach(n => {
      const emp = nomeEmp(n.empresa);
      if (!empMap[emp]) empMap[emp] = { total:0, ativas:0, criticas:0, atencao:0, ok:0, valorTotal:0, custoTotal:0, taxas:0, icms:0, multas:0, juros:0 };
      empMap[emp].total++;
      if (n.status !== "Desembaraçada" && n.status !== "Recusada") empMap[emp].ativas++;
      if (num(n.qtdeDias) >= 60) empMap[emp].criticas++;
      else if (num(n.qtdeDias) >= 25) empMap[emp].atencao++;
      else empMap[emp].ok++;
      empMap[emp].valorTotal += num(n.valor);
      const custo = num(n.taxaReanalise)+num(n.taxaDesembaraco)+num(n.icmsAntecipado)+num(n.multa10pct||0)+num(n.multa)+num(n.juros);
      empMap[emp].custoTotal += custo;
      empMap[emp].taxas += num(n.taxaReanalise)+num(n.taxaDesembaraco);
      empMap[emp].icms += num(n.icmsAntecipado);
      empMap[emp].multas += num(n.multa);
      empMap[emp].juros += num(n.juros);
    });
    const resumoEmp = Object.entries(empMap).map(([emp, d]) => ({
      "Empresa": emp,
      "Total Notas": d.total,
      "Notas Ativas": d.ativas,
      "Críticas (60+d)": d.criticas,
      "Atenção (25-59d)": d.atencao,
      "OK (0-24d)": d.ok,
      "% Críticas": d.total > 0 ? ((d.criticas/d.total)*100).toFixed(1)+"%" : "0%",
      "Valor Total NF (R$)": d.valorTotal,
      "Taxas (R$)": d.taxas,
      "ICMS Antecipado (R$)": d.icms,
      "Multas (R$)": d.multas,
      "Juros (R$)": d.juros,
      "Custo Total (R$)": d.custoTotal,
    }));
    const ws2 = XLSX.utils.json_to_sheet(resumoEmp);
    ws2["!cols"] = [{wch:24},{wch:13},{wch:13},{wch:14},{wch:16},{wch:12},{wch:12},{wch:18},{wch:14},{wch:20},{wch:14},{wch:12},{wch:16}];
    XLSX.utils.book_append_sheet(wb, ws2, "Resumo por Empresa");

    // ---- ABA 3: ANÁLISE DE PRAZOS ----
    const faixas = [
      { label: "0-15 dias", min:0, max:15 },
      { label: "16-30 dias", min:16, max:30 },
      { label: "31-45 dias", min:31, max:45 },
      { label: "46-60 dias", min:46, max:60 },
      { label: "61-90 dias", min:61, max:90 },
      { label: "91-120 dias", min:91, max:120 },
      { label: "121-180 dias", min:121, max:180 },
      { label: "180+ dias", min:181, max:99999 },
    ];
    const analisePrazos = faixas.map(f => {
      const notas = filtradas.filter(n => num(n.qtdeDias) >= f.min && num(n.qtdeDias) <= f.max);
      const custos = notas.reduce((s,n) => s+num(n.taxaReanalise)+num(n.taxaDesembaraco)+num(n.icmsAntecipado)+num(n.multa10pct||0)+num(n.multa)+num(n.juros), 0);
      return {
        "Faixa de Prazo": f.label,
        "Qtde Notas": notas.length,
        "% do Total": filtradas.length > 0 ? ((notas.length/filtradas.length)*100).toFixed(1)+"%" : "0%",
        "Valor Total NF (R$)": notas.reduce((s,n) => s+num(n.valor), 0),
        "Custo Acumulado (R$)": custos,
        "Risco": f.min >= 60 ? "ALTO" : f.min >= 25 ? "MÉDIO" : "BAIXO",
      };
    });
    const ws3 = XLSX.utils.json_to_sheet(analisePrazos);
    ws3["!cols"] = [{wch:16},{wch:12},{wch:12},{wch:18},{wch:20},{wch:10}];
    XLSX.utils.book_append_sheet(wb, ws3, "Análise de Prazos");

    // ---- ABA 4: CUSTOS DETALHADOS ----
    const totalNotas = filtradas.length;
    const totalValor = filtradas.reduce((s,n) => s+num(n.valor), 0);
    const totalTaxas = filtradas.reduce((s,n) => s+num(n.taxaReanalise)+num(n.taxaDesembaraco), 0);
    const totalIcms = filtradas.reduce((s,n) => s+num(n.icmsAntecipado), 0);
    const totalMultas = filtradas.reduce((s,n) => s+num(n.multa), 0);
    const totalJuros = filtradas.reduce((s,n) => s+num(n.juros), 0);
    const totalCustos = totalTaxas+totalIcms+totalMultas+totalJuros;
    const custosDetalhados = [
      { "Categoria": "Taxas (Reanálise + Desembaraço)", "Valor (R$)": totalTaxas, "% do Custo Total": totalCustos > 0 ? ((totalTaxas/totalCustos)*100).toFixed(1)+"%" : "0%", "Descrição": "Taxas cobradas pela SEFAZ/AM" },
      { "Categoria": "ICMS Antecipado", "Valor (R$)": totalIcms, "% do Custo Total": totalCustos > 0 ? ((totalIcms/totalCustos)*100).toFixed(1)+"%" : "0%", "Descrição": "Recolhimento antecipado para notas abaixo de R$25k" },
      { "Categoria": "Multas", "Valor (R$)": totalMultas, "% do Custo Total": totalCustos > 0 ? ((totalMultas/totalCustos)*100).toFixed(1)+"%" : "0%", "Descrição": "Multa 10% para notas acima de R$25k e multa por dia após 60 dias" },
      { "Categoria": "Juros", "Valor (R$)": totalJuros, "% do Custo Total": totalCustos > 0 ? ((totalJuros/totalCustos)*100).toFixed(1)+"%" : "0%", "Descrição": "Juros acumulados por atraso" },
      { "Categoria": "TOTAL CUSTOS", "Valor (R$)": totalCustos, "% do Custo Total": "100%", "Descrição": "" },
      { "Categoria": "", "Valor (R$)": "", "% do Custo Total": "", "Descrição": "" },
      { "Categoria": "INDICADORES GERAIS", "Valor (R$)": "", "% do Custo Total": "", "Descrição": "" },
      { "Categoria": "Total de notas no período", "Valor (R$)": totalNotas, "% do Custo Total": "", "Descrição": "" },
      { "Categoria": "Valor total das NF-e", "Valor (R$)": totalValor, "% do Custo Total": "", "Descrição": "" },
      { "Categoria": "Custo médio por nota", "Valor (R$)": totalNotas > 0 ? (totalCustos/totalNotas).toFixed(2) : 0, "% do Custo Total": "", "Descrição": "" },
      { "Categoria": "Notas críticas (60+d)", "Valor (R$)": filtradas.filter(n => num(n.qtdeDias) >= 60).length, "% do Custo Total": "", "Descrição": "Risco alto de multa adicional" },
      { "Categoria": "Notas acima R$25k", "Valor (R$)": filtradas.filter(n => num(n.valor) >= 25000).length, "% do Custo Total": "", "Descrição": "Requerem e-mail à SEFAZ + multa 10%" },
    ];
    const ws4 = XLSX.utils.json_to_sheet(custosDetalhados);
    ws4["!cols"] = [{wch:36},{wch:16},{wch:18},{wch:50}];
    XLSX.utils.book_append_sheet(wb, ws4, "Custos Detalhados");

    // ---- ABA 5: RANKING FORNECEDORES ----
    const fornMap = {};
    filtradas.forEach(n => {
      const k = n.razaoSocial || n.fornecedor || "Desconhecido";
      if (!fornMap[k]) fornMap[k] = { qtde:0, valor:0, custos:0, diasMax:0 };
      fornMap[k].qtde++;
      fornMap[k].valor += num(n.valor);
      fornMap[k].custos += num(n.taxaReanalise)+num(n.taxaDesembaraco)+num(n.icmsAntecipado)+num(n.multa10pct||0)+num(n.multa)+num(n.juros);
      if (num(n.qtdeDias) > fornMap[k].diasMax) fornMap[k].diasMax = num(n.qtdeDias);
    });
    const rankingForn = Object.entries(fornMap)
      .sort((a,b) => b[1].custos - a[1].custos)
      .map(([nome, d], i) => ({
        "Ranking": i+1,
        "Fornecedor / Razão Social": nome,
        "Qtde Notas": d.qtde,
        "Valor Total NF (R$)": d.valor,
        "Custo Acumulado (R$)": d.custos,
        "Maior Prazo (dias)": d.diasMax,
        "Risco": d.diasMax >= 60 ? "ALTO" : d.diasMax >= 25 ? "MÉDIO" : "BAIXO",
      }));
    const ws5 = XLSX.utils.json_to_sheet(rankingForn);
    ws5["!cols"] = [{wch:10},{wch:40},{wch:12},{wch:18},{wch:20},{wch:18},{wch:10}];
    XLSX.utils.book_append_sheet(wb, ws5, "Ranking Fornecedores");

    const hoje = new Date().toLocaleDateString("pt-BR").replace(/\//g,"-");
    XLSX.writeFile(wb, `DTE-Enerwatt-${periodo}-${hoje}.xlsx`);
  } catch(e) {
    alert("Erro ao gerar Excel: " + e.message);
  }
}

// ============================================================
// EXPORTAÇÃO PDF COM GRÁFICOS — jsPDF + Canvas nativo
// Layout fixo e controlado por coordenadas absolutas
// ============================================================
function gerarGraficoBarras(labels, values, colors, titulo, w, h) {
  const canvas = document.createElement("canvas");
  canvas.width = w * 2; canvas.height = h * 2;
  const ctx = canvas.getContext("2d");
  ctx.scale(2, 2);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, w, h);

  const pad = { top: 32, right: 14, bottom: 48, left: 58 };
  const cw = w - pad.left - pad.right;
  const ch = h - pad.top - pad.bottom;
  const maxVal = Math.max(...values, 1);
  const barW = Math.max(8, Math.min(40, (cw / Math.max(values.length, 1)) * 0.65));
  const gap = (cw / values.length) * 0.4;

  // Título
  ctx.fillStyle = "#1a4a4a";
  ctx.font = "bold 9px Arial";
  ctx.textAlign = "left";
  ctx.fillText(titulo, pad.left, 18);

  // Grid linhas
  ctx.strokeStyle = "#e8f0f0";
  ctx.lineWidth = 0.5;
  for (let i = 0; i <= 4; i++) {
    const y = pad.top + (ch / 4) * i;
    ctx.beginPath(); ctx.moveTo(pad.left, y); ctx.lineTo(pad.left + cw, y); ctx.stroke();
    const val = maxVal * (1 - i / 4);
    ctx.fillStyle = "#8aacac";
    ctx.font = "6px Arial";
    ctx.textAlign = "right";
    ctx.fillText(val >= 1000 ? (val/1000).toFixed(0)+"k" : val.toFixed(0), pad.left - 4, y + 2);
  }

  // Barras
  values.forEach((v, i) => {
    const x = pad.left + (cw / values.length) * i + gap / 2;
    const bh = (v / maxVal) * ch;
    const y = pad.top + ch - bh;
    ctx.fillStyle = colors[i % colors.length];
    ctx.beginPath();
    ctx.roundRect ? ctx.roundRect(x, y, barW, bh, [3, 3, 0, 0]) : ctx.rect(x, y, barW, bh);
    ctx.fill();
    // Valor em cima
    ctx.fillStyle = "#1a4a4a";
    ctx.font = "bold 6px Arial";
    ctx.textAlign = "center";
    if (v > 0) ctx.fillText(v >= 1000 ? (v/1000).toFixed(1)+"k" : v.toString(), x + barW/2, y - 3);
    // Label embaixo
    ctx.fillStyle = "#5a7a7a";
    ctx.font = "6px Arial";
    const lbl = labels[i].length > 12 ? labels[i].slice(0,11)+"…" : labels[i];
    ctx.fillText(lbl, x + barW/2, pad.top + ch + 12);
  });

  // Eixo X
  ctx.strokeStyle = "#c8dcdc";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(pad.left, pad.top + ch);
  ctx.lineTo(pad.left + cw, pad.top + ch);
  ctx.stroke();

  return canvas.toDataURL("image/png");
}

function gerarGraficoPizza(labels, values, colors, titulo, w, h) {
  const canvas = document.createElement("canvas");
  canvas.width = w * 2; canvas.height = h * 2;
  const ctx = canvas.getContext("2d");
  ctx.scale(2, 2);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, w, h);

  ctx.fillStyle = "#1a4a4a";
  ctx.font = "bold 9px Arial";
  ctx.textAlign = "left";
  ctx.fillText(titulo, 12, 18);

  const total = values.reduce((a, b) => a + b, 0);
  if (total === 0) return canvas.toDataURL("image/png");

  const cx = w * 0.38, cy = h * 0.54, r = Math.min(w, h) * 0.33;
  let startAngle = -Math.PI / 2;

  values.forEach((v, i) => {
    const slice = (v / total) * 2 * Math.PI;
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, r, startAngle, startAngle + slice);
    ctx.closePath();
    ctx.fillStyle = colors[i % colors.length];
    ctx.fill();
    ctx.strokeStyle = "#fff";
    ctx.lineWidth = 1.5;
    ctx.stroke();
    startAngle += slice;
  });

  // Buraco central (donut)
  ctx.beginPath();
  ctx.arc(cx, cy, r * 0.45, 0, 2 * Math.PI);
  ctx.fillStyle = "#ffffff";
  ctx.fill();

  // % total no centro
  ctx.fillStyle = "#1a4a4a";
  ctx.font = "bold 8px Arial";
  ctx.textAlign = "center";
  ctx.fillText("Total", cx, cy - 3);
  ctx.font = "6px Arial";
  ctx.fillStyle = "#5a7a7a";
  ctx.fillText(total >= 1000 ? "R$"+(total/1000).toFixed(0)+"k" : total.toString(), cx, cy + 7);

  // Legenda lateral
  const lx = w * 0.72, ly0 = h * 0.22;
  labels.forEach((lbl, i) => {
    const ly = ly0 + i * 16;
    ctx.fillStyle = colors[i % colors.length];
    ctx.fillRect(lx, ly - 5, 9, 9);
    ctx.fillStyle = "#1a4a4a";
    ctx.font = "6.5px Arial";
    ctx.textAlign = "left";
    const pct = total > 0 ? ((values[i] / total) * 100).toFixed(1) : "0";
    const lblShort = lbl.length > 14 ? lbl.slice(0, 13) + "…" : lbl;
    ctx.fillText(`${lblShort}`, lx + 12, ly + 2);
    ctx.fillStyle = "#E8450A";
    ctx.font = "bold 6px Arial";
    ctx.fillText(`${pct}%`, lx + 12, ly + 10);
  });

  return canvas.toDataURL("image/png");
}

function exportarPDF(filtradas, empresas, periodo) {
  try {
    const { jsPDF } = window.jspdf;
    if (!jsPDF) { alert("Biblioteca PDF não carregada. Recarregue a página."); return; }

    const num = (v) => Number(v || 0);
    const fmtM = (v) => `R$ ${num(v).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`;
    const nomeEmp = (id) => { const e = empresas.find(x => x.id === id); return e ? e.nome.split(" - ").pop() : id; };
    const hoje = new Date().toLocaleDateString("pt-BR");
    const hojeFile = hoje.replace(/\//g, "-");

    const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
    // A4 portrait: 210 x 297 mm
    const PW = 210, PH = 297;
    const PETROL = [26, 74, 74];
    const LARANJA = [232, 69, 10];
    const CINZA = [240, 246, 246];
    const HEADER_H = 18; // altura do header
    const FOOTER_Y = 287; // posição do rodapé
    const MARGIN = 10; // margem lateral
    const CONTENT_W = PW - MARGIN * 2; // 190mm

    const drawHeader = () => {
      doc.setFillColor(...PETROL);
      doc.rect(0, 0, PW, HEADER_H, "F");
      doc.setFillColor(...LARANJA);
      doc.rect(0, HEADER_H, PW, 1.5, "F");
      doc.setTextColor(255, 255, 255);
      doc.setFont("helvetica", "bold"); doc.setFontSize(10);
      doc.text("ENERWATT — Gestão DTE/AM", MARGIN, 11);
      doc.setFont("helvetica", "normal"); doc.setFontSize(7.5);
      doc.text(`Relatório ${periodo.toUpperCase()} · ${hoje}`, PW - MARGIN, 11, { align: "right" });
    };

    const drawFooter = (pg, total) => {
      doc.setFillColor(...CINZA);
      doc.rect(0, FOOTER_Y, PW, 10, "F");
      doc.setTextColor(120, 150, 150);
      doc.setFont("helvetica", "normal"); doc.setFontSize(6.5);
      doc.text(`Enerwatt Engenharia — Sistema DTE/AM — Gerado em ${hoje} — Pág ${pg}/${total}`, PW / 2, FOOTER_Y + 6, { align: "center" });
    };

    const sectionTitle = (y, texto) => {
      doc.setFillColor(...CINZA);
      doc.rect(MARGIN, y, CONTENT_W, 9, "F");
      doc.setFillColor(...LARANJA);
      doc.rect(MARGIN, y, 2.5, 9, "F");
      doc.setTextColor(...PETROL);
      doc.setFont("helvetica", "bold"); doc.setFontSize(8.5);
      doc.text(texto, MARGIN + 6, y + 6.2);
      return y + 13; // retorna y após o título
    };

    // =============================================
    // CALCULAR DADOS
    // =============================================
    const totalNotas = filtradas.length;
    const ativas = filtradas.filter(n => !["Desembaraçada","Recusada"].includes(n.status));
    const criticas = filtradas.filter(n => num(n.qtdeDias) >= 60).length;
    const atencao = filtradas.filter(n => num(n.qtdeDias) >= 25 && num(n.qtdeDias) < 60).length;
    const totalTaxas = ativas.reduce((s,n) => s + num(n.taxaReanalise) + num(n.taxaDesembaraco), 0);
    const totalIcms = ativas.reduce((s,n) => s + num(n.icmsAntecipado), 0);
    const totalMultas = ativas.reduce((s,n) => s + num(n.multa), 0);
    const totalJuros = ativas.reduce((s,n) => s + num(n.juros), 0);
    const totalCustos = totalTaxas + totalIcms + totalMultas + totalJuros;

    // =============================================
    // PÁGINA 1: KPIs + Gráfico Prazos + Gráfico Custos
    // =============================================
    drawHeader();

    // -- Subtítulo
    let y = HEADER_H + 2;
    doc.setFillColor(250, 252, 252);
    doc.rect(0, y, PW, 11, "F");
    doc.setTextColor(...PETROL); doc.setFont("helvetica","bold"); doc.setFontSize(11);
    doc.text("Indicadores de Gestão DTE/AM", MARGIN, y + 8);
    y += 15;

    // -- KPIs (4 cards lado a lado, 44mm cada)
    const kpis = [
      { label: "Total Notas", val: String(totalNotas), cor: PETROL },
      { label: "Críticas 60+d", val: String(criticas), cor: [192, 57, 43] },
      { label: "Atenção 25-59d", val: String(atencao), cor: [200, 130, 0] },
      { label: "Custos Ativos", val: fmtM(totalCustos), cor: LARANJA },
    ];
    const kW = 45, kH = 18, kGap = 2.5;
    kpis.forEach((k, i) => {
      const x = MARGIN + i * (kW + kGap);
      doc.setFillColor(255, 255, 255);
      doc.roundedRect(x, y, kW, kH, 1.5, 1.5, "F");
      doc.setFillColor(...k.cor);
      doc.rect(x, y, 2.5, kH, "F");
      doc.setTextColor(...k.cor);
      doc.setFont("helvetica","bold");
      // valor menor se for moeda para caber
      doc.setFontSize(k.val.length > 8 ? 8 : 13);
      doc.text(k.val, x + 5, y + (k.val.length > 8 ? 8 : 11));
      doc.setFont("helvetica","normal"); doc.setFontSize(6.5);
      doc.setTextColor(90, 110, 110);
      doc.text(k.label, x + 5, y + 15);
    });
    y += kH + 6;

    // -- Gráfico 1: Faixas de prazo (60mm altura)
    y = sectionTitle(y, "Distribuição de Notas por Faixa de Prazo");
    const faixasLabels = ["0-15d","16-30d","31-45d","46-60d","61-90d","91-120d","121+d"];
    const faixasMins  = [0, 16, 31, 46, 61, 91, 121];
    const faixasMaxs  = [15, 30, 45, 60, 90, 120, 9999];
    const faixasVals  = faixasMins.map((min, i) =>
      filtradas.filter(n => num(n.qtdeDias) >= min && num(n.qtdeDias) <= faixasMaxs[i]).length
    );
    const coresFaixas = ["#4db8b8","#4db8b8","#f0a500","#f0a500","#E8450A","#c0392b","#7b0000"];
    const G1_H = 60; // altura em mm no PDF
    const g1img = gerarGraficoBarras(faixasLabels, faixasVals, coresFaixas, "", 380, 120);
    doc.setFillColor(255,255,255);
    doc.rect(MARGIN, y, CONTENT_W, G1_H, "F");
    doc.addImage(g1img, "PNG", MARGIN, y, CONTENT_W, G1_H);
    y += G1_H + 5;

    // -- Gráfico 2: Pizza custos (55mm altura)
    y = sectionTitle(y, "Composição dos Custos Registrados (Notas Ativas)");
    const G2_H = 58;
    const custoLabels = ["Taxas","ICMS Antecipado","Multas","Juros"];
    const custoVals = [totalTaxas, totalIcms, totalMultas, totalJuros];
    const coresCustos = ["#1a4a4a","#4db8b8","#E8450A","#f0a500"];
    const g2img = gerarGraficoPizza(custoLabels, custoVals, coresCustos, "", 380, 116);
    doc.setFillColor(255,255,255);
    doc.rect(MARGIN, y, CONTENT_W, G2_H, "F");
    doc.addImage(g2img, "PNG", MARGIN, y, CONTENT_W, G2_H);
    y += G2_H + 5;

    // -- Tabela resumo custos (4 colunas)
    const colW4 = CONTENT_W / 4;
    const custoCols2 = [
      { label: "Taxas", val: fmtM(totalTaxas) },
      { label: "ICMS Antecipado", val: fmtM(totalIcms) },
      { label: "Multas", val: fmtM(totalMultas) },
      { label: "Juros + Total", val: fmtM(totalCustos) },
    ];
    doc.setFillColor(...CINZA);
    doc.rect(MARGIN, y, CONTENT_W, 16, "F");
    custoCols2.forEach((c, i) => {
      const x = MARGIN + i * colW4 + 3;
      doc.setFont("helvetica","normal"); doc.setFontSize(6.5); doc.setTextColor(100,130,130);
      doc.text(c.label, x, y + 6);
      doc.setFont("helvetica","bold"); doc.setFontSize(8);
      doc.setTextColor(i === 3 ? 232 : 26, i === 3 ? 69 : 74, i === 3 ? 10 : 74);
      doc.text(c.val, x, y + 13);
    });

    // =============================================
    // PÁGINA 2: Gráficos por Empresa + Ranking
    // =============================================
    doc.addPage();
    drawHeader();
    y = HEADER_H + 5;

    // Empresa map
    const empMap = {};
    filtradas.forEach(n => {
      const nm = nomeEmp(n.empresa);
      if (!empMap[nm]) empMap[nm] = { qtde:0, criticas:0, custos:0 };
      empMap[nm].qtde++;
      if (num(n.qtdeDias) >= 60) empMap[nm].criticas++;
      empMap[nm].custos += num(n.taxaReanalise)+num(n.taxaDesembaraco)+num(n.icmsAntecipado)+num(n.multa10pct||0)+num(n.multa)+num(n.juros);
    });
    const empNomes = Object.keys(empMap);
    const empQtdes = empNomes.map(e => empMap[e].qtde);
    const empCriticas = empNomes.map(e => empMap[e].criticas);
    const empCustos = empNomes.map(e => empMap[e].custos);

    // Gráfico por empresa — quantidade (50mm)
    y = sectionTitle(y, "Notas por Empresa — Quantidade Total");
    const g3img = gerarGraficoBarras(empNomes, empQtdes, ["#1a4a4a","#4db8b8","#E8450A"], "", 380, 100);
    doc.setFillColor(255,255,255); doc.rect(MARGIN, y, CONTENT_W, 50, "F");
    doc.addImage(g3img, "PNG", MARGIN, y, CONTENT_W, 50);
    y += 55;

    // Gráfico por empresa — críticas (50mm)
    y = sectionTitle(y, "Notas Críticas (60+ dias) por Empresa");
    const g4img = gerarGraficoBarras(empNomes, empCriticas, ["#c0392b","#E8450A","#f0a500"], "", 380, 100);
    doc.setFillColor(255,255,255); doc.rect(MARGIN, y, CONTENT_W, 50, "F");
    doc.addImage(g4img, "PNG", MARGIN, y, CONTENT_W, 50);
    y += 55;

    // Ranking fornecedores
    y = sectionTitle(y, "Top 10 Fornecedores por Custo Acumulado");

    const theadCols = [MARGIN, MARGIN+68, MARGIN+90, MARGIN+112, MARGIN+144, MARGIN+170];
    doc.setFillColor(...PETROL);
    doc.rect(MARGIN, y-1, CONTENT_W, 8, "F");
    doc.setTextColor(255,255,255); doc.setFont("helvetica","bold"); doc.setFontSize(6.5);
    ["#","Razão Social","Qtde","Prazo Máx","Custo Acum.","Risco"].forEach((h,i) => doc.text(h, theadCols[i], y+5));
    y += 10;

    const fornMap = {};
    filtradas.forEach(n => {
      const k = n.razaoSocial || n.fornecedor || "Desconhecido";
      if (!fornMap[k]) fornMap[k] = { qtde:0, custos:0, diasMax:0 };
      fornMap[k].qtde++;
      fornMap[k].custos += num(n.taxaReanalise)+num(n.taxaDesembaraco)+num(n.icmsAntecipado)+num(n.multa10pct||0)+num(n.multa)+num(n.juros);
      if (num(n.qtdeDias) > fornMap[k].diasMax) fornMap[k].diasMax = num(n.qtdeDias);
    });
    Object.entries(fornMap).sort((a,b) => b[1].custos - a[1].custos).slice(0,10).forEach(([nome, d], i) => {
      if (i % 2 === 0) { doc.setFillColor(242,250,250); doc.rect(MARGIN, y-3, CONTENT_W, 7, "F"); }
      const risco = d.diasMax >= 60 ? "ALTO" : d.diasMax >= 25 ? "MÉDIO" : "BAIXO";
      const rc = d.diasMax >= 60 ? [192,57,43] : d.diasMax >= 25 ? [180,110,0] : [30,110,70];
      doc.setFont("helvetica","normal"); doc.setFontSize(6.5); doc.setTextColor(50,70,70);
      doc.text(String(i+1), theadCols[0], y+2);
      doc.text(nome.slice(0,34), theadCols[1], y+2);
      doc.text(String(d.qtde), theadCols[2], y+2);
      doc.text(`${d.diasMax}d`, theadCols[3], y+2);
      doc.text(fmtM(d.custos), theadCols[4], y+2);
      doc.setTextColor(...rc); doc.setFont("helvetica","bold");
      doc.text(risco, theadCols[5], y+2);
      y += 8;
    });

    // =============================================
    // PÁGINA 3+: Listagem completa
    // =============================================
    doc.addPage();
    drawHeader();
    let ly = HEADER_H + 5;
    ly = sectionTitle(ly, "Listagem Completa de Notas");

    const lCols = [MARGIN, MARGIN+55, MARGIN+80, MARGIN+94, MARGIN+108, MARGIN+118, MARGIN+145, MARGIN+170];
    const lHeads = ["Razão Social","NF","CFOP","Emissão","Dias","Status","Custo Total"];

    const drawTableHeader = () => {
      doc.setFillColor(...PETROL);
      doc.rect(MARGIN, ly-2, CONTENT_W, 8, "F");
      doc.setTextColor(255,255,255); doc.setFont("helvetica","bold"); doc.setFontSize(6.5);
      doc.text("Empresa", lCols[0], ly+4);
      lHeads.forEach((h,i) => doc.text(h, lCols[i+1], ly+4));
      ly += 10;
    };
    drawTableHeader();

    filtradas.forEach((n, i) => {
      if (ly > FOOTER_Y - 10) {
        drawFooter(doc.internal.getNumberOfPages(), "?");
        doc.addPage();
        drawHeader();
        ly = HEADER_H + 5;
        drawTableHeader();
      }
      if (i % 2 === 0) { doc.setFillColor(244,250,250); doc.rect(MARGIN, ly-3, CONTENT_W, 7, "F"); }
      const custo = num(n.taxaReanalise)+num(n.taxaDesembaraco)+num(n.icmsAntecipado)+num(n.multa10pct||0)+num(n.multa)+num(n.juros);
      const isCrit = num(n.qtdeDias) >= 60;
      doc.setFont("helvetica","normal"); doc.setFontSize(6.5); doc.setTextColor(50,70,70);
      doc.text(nomeEmp(n.empresa).slice(0,18), lCols[0], ly+2);
      doc.text((n.razaoSocial||"").slice(0,26), lCols[1], ly+2);
      doc.text(n.numNota||"", lCols[2], ly+2);
      doc.text(n.cfop||"", lCols[3], ly+2);
      doc.text(n.dtEmissao||"", lCols[4], ly+2);
      doc.setFont("helvetica", isCrit ? "bold" : "normal");
      doc.setTextColor(...(isCrit ? [192,57,43] : [50,70,70]));
      doc.text(`${n.qtdeDias||0}d`, lCols[5], ly+2);
      doc.setFont("helvetica","normal"); doc.setTextColor(50,70,70);
      doc.text((n.status||"").slice(0,18), lCols[6], ly+2);
      doc.setFont("helvetica", custo > 0 ? "bold" : "normal");
      doc.setTextColor(custo > 0 ? 232 : 50, custo > 0 ? 69 : 70, custo > 0 ? 10 : 70);
      doc.text(fmtM(custo), lCols[7], ly+2);
      ly += 7;
    });

    // Rodapés em todas as páginas
    const totalPgs = doc.internal.getNumberOfPages();
    for (let p = 1; p <= totalPgs; p++) {
      doc.setPage(p);
      drawFooter(p, totalPgs);
    }

    doc.save(`DTE-Enerwatt-${periodo}-${hojeFile}.pdf`);
  } catch(e) {
    alert("Erro ao gerar PDF: " + e.message);
    console.error(e);
  }
}

// ============================================================
// DADOS INICIAIS — importados dos CSVs das 3 empresas
// ============================================================
const EMPRESAS_INICIAIS = [
  { id: "filial02", nome: "FILIAL 02 - MANAUS", inscricao: "04.235.429-3", cnpj: "07.791.042/0002-18", ativa: true },
  { id: "filial13", nome: "FILIAL 13 - MLA", inscricao: "04.235.429-X", cnpj: "07.791.042/0013-00", ativa: true },
  { id: "linhasnorte", nome: "LINHAS DO NORTE", inscricao: "04.235.429-Y", cnpj: "00.000.000/0001-00", ativa: true },
];
// Alias global para compatibilidade (será sobrescrito pelo estado do App via prop)
let EMPRESAS = EMPRESAS_INICIAIS;

const FINALIDADES = ["Uso/Consumo", "Industrialização", "Revenda", "Remessa/Transferência", "Imobilizado", "Não Identificado"];
const STATUS_LIST = ["Identificada", "Em Identificação", "Em Reanálise", "Aguardando Pagamento", "Aguarda Email SEFAZ", "Desembaraço Solicitado", "Desembaraçada", "Selada", "Recusada", "Postergada"];

function parseValor(v) {
  if (!v) return 0;
  return parseFloat(v.replace(/\./g, "").replace(",", ".")) || 0;
}

function parseDateBR(d) {
  if (!d || d === "-") return null;
  const [dia, mes, ano] = d.split("/");
  return new Date(`${ano}-${mes}-${dia}`);
}

function diffDias(dataEmissao) {
  if (!dataEmissao) return 0;
  const hoje = new Date("2026-03-06");
  return Math.floor((hoje - dataEmissao) / (1000 * 60 * 60 * 24));
}

const NOTAS_INICIAIS = [
  // FILIAL 02 - MANAUS
  { id: "1", empresa: "filial02", fornecedor: "02.677.045/0002-01", razaoSocial: "HORUS TELECOMUNICACOES LTDA", numNota: "184028", cfop: "6102", dtEmissao: "07/01/2026", dtApresentacao: "-", chave: "52260102677045000201552110001840281230738747", valor: 3623.50, qtdeDias: 58, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "2", empresa: "filial02", fornecedor: "02.677.045/0002-01", razaoSocial: "HORUS TELECOMUNICACOES LTDA", numNota: "184154", cfop: "6102", dtEmissao: "12/01/2026", dtApresentacao: "-", chave: "52260102677045000201552110001841541263066581", valor: 6197.64, qtdeDias: 53, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "3", empresa: "filial02", fornecedor: "07.791.042/0001-37", razaoSocial: "ENERWATT ENGENHARIA LTDA", numNota: "5327", cfop: "6554", dtEmissao: "12/01/2026", dtApresentacao: "-", chave: "52260107791042000137550010000053271163320077", valor: 240000.00, qtdeDias: 53, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "4", empresa: "filial02", fornecedor: "07.791.042/0001-37", razaoSocial: "ENERWATT ENGENHARIA LTDA", numNota: "5332", cfop: "6554", dtEmissao: "19/01/2026", dtApresentacao: "-", chave: "52260107791042000137550010000053321351857147", valor: 97000.00, qtdeDias: 46, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "5", empresa: "filial02", fornecedor: "26.502.220/0001-07", razaoSocial: "Engecomp Consultoria e Locacao de Sistemas Ltda", numNota: "2064", cfop: "6908", dtEmissao: "19/01/2026", dtApresentacao: "-", chave: "35260126502220000107550010000020641140106616", valor: 1200.00, qtdeDias: 46, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "6", empresa: "filial02", fornecedor: "01.816.875/0001-29", razaoSocial: "AJEL MATERIAIS ELETRICOS LTDA", numNota: "1045837", cfop: "6110", dtEmissao: "30/01/2026", dtApresentacao: "-", chave: "52260101816875000129550010010458371541541030", valor: 1190.05, qtdeDias: 35, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "7", empresa: "filial02", fornecedor: "35.784.562/0001-58", razaoSocial: "ADR COMERCIO DE EQUIPAMENTOS DE INFORMATICA EIRELI", numNota: "13254", cfop: "6102", dtEmissao: "09/02/2026", dtApresentacao: "-", chave: "35260235784562000158550010000132541666574003", valor: 6130.00, qtdeDias: 25, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "8", empresa: "filial02", fornecedor: "13.087.023/0001-27", razaoSocial: "BRASFORMER PRODUTOS ELETRICOS LTDA", numNota: "10154", cfop: "6109", dtEmissao: "20/02/2026", dtApresentacao: "-", chave: "35260213087023000127550010000101541201141316", valor: 3745.10, qtdeDias: 14, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  // FILIAL 13 - MLA
  { id: "9", empresa: "filial13", fornecedor: "07.791.042/0007-22", razaoSocial: "ENERWATT ENGENHARIA, INDUSTRIA E COMERCIO - EIRELI", numNota: "1702", cfop: "6151", dtEmissao: "30/01/2026", dtApresentacao: "-", chave: "35260107791042000722550010000017021171724335", valor: 4272.61, qtdeDias: 35, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "10", empresa: "filial13", fornecedor: "07.791.042/0001-37", razaoSocial: "ENERWATT ENGENHARIA LTDA", numNota: "5364", cfop: "6554", dtEmissao: "16/02/2026", dtApresentacao: "-", chave: "52260207791042000137550010000053641305471689", valor: 4509.90, qtdeDias: 18, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "11", empresa: "filial13", fornecedor: "02.341.470/0001-44", razaoSocial: "Boa Vista Energia S/A", numNota: "1017", cfop: "6915", dtEmissao: "20/02/2026", dtApresentacao: "-", chave: "14260202341470000144550020000010171311612194", valor: 99610.55, qtdeDias: 14, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "12", empresa: "filial13", fornecedor: "02.341.470/0001-44", razaoSocial: "Boa Vista Energia S/A", numNota: "1016", cfop: "6915", dtEmissao: "20/02/2026", dtApresentacao: "-", chave: "14260202341470000144550020000010161341856914", valor: 172410.30, qtdeDias: 14, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  // LINHAS DO NORTE
  { id: "13", empresa: "linhasnorte", fornecedor: "10.159.093/0002-36", razaoSocial: "VIMEZER FORN DE SERV LTDA", numNota: "886481", cfop: "6403", dtEmissao: "17/01/2026", dtApresentacao: "-", chave: "14260110159093000236550010008864811598662666", valor: 9030.00, qtdeDias: 48, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "14", empresa: "linhasnorte", fornecedor: "01.200.900/0001-45", razaoSocial: "ELEUZA AMARAL DA SILVA", numNota: "939", cfop: "6103", dtEmissao: "21/01/2026", dtApresentacao: "-", chave: "14260101200900000145550010000009391300001762", valor: 7357.00, qtdeDias: 44, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "15", empresa: "linhasnorte", fornecedor: "02.905.133/0001-32", razaoSocial: "CABEXPRESS IND.COM.DE CABOS ELET.", numNota: "43738", cfop: "6101", dtEmissao: "21/01/2026", dtApresentacao: "-", chave: "35260102905133000132550010000437381888458950", valor: 6762.24, qtdeDias: 44, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "16", empresa: "linhasnorte", fornecedor: "09.296.337/0001-62", razaoSocial: "C. COMERCIO CONSTRUCAO SERVICOS LTDA", numNota: "2816", cfop: "6102", dtEmissao: "21/01/2026", dtApresentacao: "-", chave: "14260109296337000162550010000028161380042518", valor: 1600.00, qtdeDias: 44, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "17", empresa: "linhasnorte", fornecedor: "09.296.337/0001-62", razaoSocial: "C. COMERCIO CONSTRUCAO SERVICOS LTDA", numNota: "2815", cfop: "6102", dtEmissao: "21/01/2026", dtApresentacao: "-", chave: "14260109296337000162550010000028151518643526", valor: 800.00, qtdeDias: 44, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "18", empresa: "linhasnorte", fornecedor: "10.159.093/0002-36", razaoSocial: "VIMEZER FORN DE SERV LTDA", numNota: "887052", cfop: "6403", dtEmissao: "26/01/2026", dtApresentacao: "-", chave: "14260110159093000236550010008870521446156556", valor: 1210.00, qtdeDias: 39, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "19", empresa: "linhasnorte", fornecedor: "62.384.763/0001-30", razaoSocial: "CIVITELLA E CIA LTDA", numNota: "18531", cfop: "6101", dtEmissao: "26/01/2026", dtApresentacao: "-", chave: "35260162384763000130550000000185311023587016", valor: 14031.89, qtdeDias: 39, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "20", empresa: "linhasnorte", fornecedor: "10.159.093/0007-40", razaoSocial: "VIMEZER FORNC. DE SERV. LTDA", numNota: "2259", cfop: "6403", dtEmissao: "30/01/2026", dtApresentacao: "-", chave: "14260110159093000740550010000022591912085940", valor: 1055.00, qtdeDias: 35, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "21", empresa: "linhasnorte", fornecedor: "01.200.900/0001-45", razaoSocial: "ELEUZA AMARAL DA SILVA", numNota: "954", cfop: "6103", dtEmissao: "04/02/2026", dtApresentacao: "-", chave: "14260201200900000145550010000009541300001916", valor: 11023.00, qtdeDias: 30, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "22", empresa: "linhasnorte", fornecedor: "05.059.252/0001-00", razaoSocial: "MOURAO E LIRA LTDA - EPP", numNota: "27692", cfop: "6102", dtEmissao: "04/02/2026", dtApresentacao: "-", chave: "14260205059252000100550010000276921095977758", valor: 1894.60, qtdeDias: 30, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "23", empresa: "linhasnorte", fornecedor: "19.215.087/0002-23", razaoSocial: "MARTINS & SA LTDA ME", numNota: "2109", cfop: "6102", dtEmissao: "05/02/2026", dtApresentacao: "-", chave: "14260219215087000223550010000021091182559020", valor: 7134.00, qtdeDias: 29, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "24", empresa: "linhasnorte", fornecedor: "10.159.093/0001-55", razaoSocial: "VIMEZER FORN. DE SERV. LTDA", numNota: "193912", cfop: "6403", dtEmissao: "11/02/2026", dtApresentacao: "-", chave: "14260210159093000155550010001939121384724670", valor: 75.00, qtdeDias: 23, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "25", empresa: "linhasnorte", fornecedor: "27.127.974/0001-97", razaoSocial: "J. F. MOREIRA -ME", numNota: "4125", cfop: "6102", dtEmissao: "11/02/2026", dtApresentacao: "-", chave: "14260227127974000197550020000041251344287275", valor: 112.00, qtdeDias: 23, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "26", empresa: "linhasnorte", fornecedor: "27.127.974/0001-97", razaoSocial: "J. F. MOREIRA -ME", numNota: "4124", cfop: "6102", dtEmissao: "11/02/2026", dtApresentacao: "-", chave: "14260227127974000197550020000041241335350698", valor: 1073.00, qtdeDias: 23, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "27", empresa: "linhasnorte", fornecedor: "10.159.093/0007-40", razaoSocial: "VIMEZER FORNC. DE SERV. LTDA", numNota: "2614", cfop: "6403", dtEmissao: "16/02/2026", dtApresentacao: "-", chave: "14260210159093000740550010000026141073344889", valor: 1783.00, qtdeDias: 18, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "28", empresa: "linhasnorte", fornecedor: "19.215.087/0002-23", razaoSocial: "MARTINS & SA LTDA ME", numNota: "2142", cfop: "6102", dtEmissao: "16/02/2026", dtApresentacao: "-", chave: "14260219215087000223550010000021421064331406", valor: 3928.00, qtdeDias: 18, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "29", empresa: "linhasnorte", fornecedor: "01.200.900/0001-45", razaoSocial: "ELEUZA AMARAL DA SILVA", numNota: "960", cfop: "6103", dtEmissao: "18/02/2026", dtApresentacao: "-", chave: "14260201200900000145550010000009601300001972", valor: 12331.00, qtdeDias: 16, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "30", empresa: "linhasnorte", fornecedor: "01.867.060/0001-79", razaoSocial: "M. J. M. DA SILVA", numNota: "17514", cfop: "6101", dtEmissao: "18/02/2026", dtApresentacao: "-", chave: "14260201867060000179550010000175141004126555", valor: 400.00, qtdeDias: 16, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
  { id: "31", empresa: "linhasnorte", fornecedor: "27.127.974/0001-97", razaoSocial: "J. F. MOREIRA -ME", numNota: "4139", cfop: "6102", dtEmissao: "24/02/2026", dtApresentacao: "-", chave: "14260227127974000197550020000041391996657711", valor: 1960.00, qtdeDias: 10, status: "Identificada", centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0, dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", codigoCusto: "", lancamentos: [], obs: "", pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "", historico: [{ acao: "Nota importada do DTE", usuario: "Sistema", data: "06/03/2026 08:00" }] },
];

// Usuários carregados do Supabase dinamicamente

// ============================================================
// TELA DE LOGIN
// ============================================================
function TelaLogin({ onLogin, logoUrl }) {
  const [email, setEmail] = useState("");
  const [senha, setSenha] = useState("");
  const [erro, setErro] = useState("");
  const [carregando, setCarregando] = useState(false);

  const handleLogin = async () => {
    if (!email || !senha) { setErro("Preencha e-mail e senha."); return; }
    setCarregando(true);
    setErro("");
    const { data, error } = await supabase.auth.signInWithPassword({ email, password: senha });
    if (error) {
      setErro("E-mail ou senha inválidos.");
      setCarregando(false);
      return;
    }
    // Buscar perfil do usuário na tabela usuarios
    const { data: perfil } = await supabase.from("usuarios").select("*").eq("email", email).single();
    // Registrar log de acesso
    const agora = new Date().toISOString();
    if (perfil?.id) {
      await supabase.from("usuarios").update({ ultimo_acesso: agora }).eq("id", perfil.id);
      await supabase.from("logs_acesso").insert({
        usuario_id: perfil.id,
        usuario_nome: perfil.nome || email,
        usuario_email: email,
        perfil: perfil.perfil || "operador",
        acessado_em: agora
      });
    }
    onLogin({ ...data.user, nome: perfil?.nome || email, perfil: perfil?.perfil || "operador" });
    setCarregando(false);
  };

  return (
    <div className="min-h-screen flex" style={{ fontFamily: "'Segoe UI', system-ui, sans-serif", background: "white", position: "relative", overflow: "hidden" }}>
      {/* Barra laranja topo */}
      <div style={{ position: "fixed", top: 0, left: 0, right: 0, height: "4px", background: "#E8450A", zIndex: 50 }} />

      {/* LADO ESQUERDO — formulário limpo branco */}
      <div className="flex flex-col justify-center" style={{ width: "45%", minHeight: "100vh", padding: "3rem 4rem", position: "relative", zIndex: 10, background: "white" }}>
        {/* Logo */}
        <div className="mb-10">
          {logoUrl
            ? <img src={logoUrl} alt="Logo" className="h-12 object-contain" />
            : (
              <div>
                <p className="font-black text-3xl" style={{ color: "#1a4a4a", fontFamily: "Georgia, serif", fontStyle: "italic" }}>Enerwatt</p>
                <p className="text-xs tracking-widest uppercase" style={{ color: "#E8450A", marginTop: 2 }}>Engenharia</p>
              </div>
            )
          }
        </div>

        {/* Texto de boas-vindas */}
        <div className="mb-8">
          <p className="text-xs font-bold uppercase tracking-widest mb-1" style={{ color: "#E8450A" }}>DEPARTAMENTO FISCAL</p>
          <h1 className="text-3xl font-black mb-2" style={{ color: "#1a4a4a" }}>Gestão DTE/AM</h1>
          <p className="text-sm" style={{ color: "#6b8c8c" }}>Sistema de Controle de Desembaraço Extemporâneo</p>
        </div>

        {/* Formulário */}
        <div className="space-y-4" style={{ maxWidth: 360 }}>
          <div>
            <label className="text-xs font-bold uppercase tracking-wide" style={{ color: "#1a4a4a" }}>E-mail</label>
            <input type="email" value={email} onChange={e => setEmail(e.target.value)}
              onKeyDown={e => e.key === "Enter" && handleLogin()}
              placeholder="seu@enerwatt.com.br"
              className="mt-1.5 w-full rounded-lg px-4 py-3 text-sm focus:outline-none"
              style={{ border: "1.5px solid #c8dede", background: "#f8fafa", color: "#1a4a4a" }} />
          </div>
          <div>
            <label className="text-xs font-bold uppercase tracking-wide" style={{ color: "#1a4a4a" }}>Senha</label>
            <input type="password" value={senha} onChange={e => setSenha(e.target.value)}
              onKeyDown={e => e.key === "Enter" && handleLogin()}
              placeholder="••••••••"
              className="mt-1.5 w-full rounded-lg px-4 py-3 text-sm focus:outline-none"
              style={{ border: "1.5px solid #c8dede", background: "#f8fafa", color: "#1a4a4a" }} />
          </div>
          {erro && <p className="text-xs font-semibold" style={{ color: "#c0392b" }}>{erro}</p>}
          <button onClick={handleLogin} disabled={carregando}
            className="w-full py-3 rounded-lg text-sm font-bold text-white disabled:opacity-60 transition-opacity hover:opacity-90"
            style={{ background: "#E8450A", marginTop: 8 }}>
            {carregando ? "Entrando..." : "Entrar →"}
          </button>
        </div>

        <p className="text-xs mt-8" style={{ color: "#aec4c4" }}>Acesso restrito — Enerwatt Engenharia</p>
      </div>

      {/* LADO DIREITO — fundo azul-teal com torre */}
      <div style={{ width: "55%", minHeight: "100vh", position: "relative", overflow: "hidden",
        background: "linear-gradient(145deg, #b8d8d8 0%, #7ecece 30%, #4db8b8 60%, #1a7a7a 100%)" }}>
        {/* Torres SVG grandes */}
        <TorresSVG opacity={0.35} corTorre="#0d3535" />
        {/* Brilho suave */}
        <div style={{ position: "absolute", inset: 0,
          background: "radial-gradient(ellipse at 60% 30%, rgba(255,255,255,0.18) 0%, transparent 55%), radial-gradient(ellipse at 30% 80%, rgba(13,53,53,0.25) 0%, transparent 50%)" }} />
        {/* Tag no canto inferior */}
        <div style={{ position: "absolute", bottom: 32, left: 32 }}>
          <p className="text-sm font-bold" style={{ color: "rgba(255,255,255,0.9)" }}>Transformar Energia</p>
          <p className="text-sm font-black" style={{ color: "white" }}>em Desenvolvimento</p>
          <div style={{ width: 40, height: 3, background: "#E8450A", borderRadius: 2, marginTop: 6 }} />
        </div>
      </div>
    </div>
  );
}

// ============================================================
// HELPERS
// ============================================================
function fmtMoeda(v) {
  return (v || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}

function getAlertas(nota) {
  const dias = nota.qtdeDias;
  const alertas = [];
  if (["Desembaraçada", "Recusada"].includes(nota.status)) return alertas;
  if (nota.valor > 25000) alertas.push({ tipo: "danger", msg: "Nota acima de R$ 25.000 — risco de multa de 10%" });
  if (dias >= 60) alertas.push({ tipo: "danger", msg: `${dias} dias — PENDÊNCIA SEFAZ! Multa diária ativa` });
  else if (dias >= 50) alertas.push({ tipo: "danger", msg: `${dias} dias — Atenção! Próximo de 60 dias (multa diária)` });
  else if (dias >= 25 && dias < 30) alertas.push({ tipo: "warning", msg: `${dias} dias — Reanalisar antes de 30 dias para evitar taxa de R$ 50` });
  if (nota.status === "Postergada" && nota.dtPostergacao) {
    const dtPost = parseDateBR(nota.dtPostergacao);
    if (dtPost) {
      const limite = new Date(dtPost);
      limite.setDate(limite.getDate() + 120);
      const hoje = new Date("2026-03-06");
      const restam = Math.floor((limite - hoje) / (1000 * 60 * 60 * 24));
      if (restam <= 30) alertas.push({ tipo: "warning", msg: `Postergação vence em ${restam} dias` });
    }
  }
  if (alertas.length === 0 && dias < 25) alertas.push({ tipo: "ok", msg: "Dentro do prazo" });
  return alertas;
}

function getProximoPasso(nota) {
  const dias = nota.qtdeDias;
  if (nota.status === "Desembaraçada") return "✅ Concluída — Desembaraçada";
  if (nota.status === "Recusada") return "❌ Concluída — Recusada no portal";
  if (nota.status === "Identificada") return "🔍 1º Passo: Realizar Reanálise — identificar CC, Finalidade e Responsável";
  if (nota.status === "Em Reanálise") {
    if (nota.noRM === false) return "📞 Acionar comprador/responsável — nota não está no RM";
    if (nota.noRM === null) return "🔍 Verificar se nota está lançada no RM";
    if (nota.valor > 25000) return "📧 Nota >R$25k — Solicitar desembaraço via EMAIL à SEFAZ com justificativa";
    return "📋 Solicitar Desembaraço — nota confirmada no RM";
  }
  if (nota.status === "Aguardando Pagamento") return "💰 Aguardando Financeiro efetuar pagamento de taxa/ICMS";
  if (nota.status === "Aguarda Email SEFAZ") return "📧 Aguardando retorno da SEFAZ ao email enviado";
  if (nota.status === "Desembaraço Solicitado") return "⏳ Solicitação em análise pela SEFAZ/AM";
  if (nota.status === "Postergada") return "🗓️ Nota postergada — monitorar prazo de 180 dias";
  return "—";
}

// ============================================================
// COMPONENTES
// ============================================================

const LOGO_URL = "data:image/png;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAA=";

function Badge({ status }) {
  const map = {
    "Identificada": "bg-gray-100 text-gray-700",
    "Em Reanálise": "bg-blue-100 text-blue-700",
    "Aguardando Pagamento": "bg-yellow-100 text-yellow-800",
    "Aguarda Email SEFAZ": "bg-purple-100 text-purple-700",
    "Desembaraço Solicitado": "bg-orange-100 text-orange-700",
    "Desembaraçada": "bg-green-100 text-green-700",
    "Recusada": "bg-red-100 text-red-700",
    "Postergada": "bg-indigo-100 text-indigo-700",
  };
  return (
    <span className={`px-2 py-0.5 rounded-full text-xs font-semibold ${map[status] || "bg-gray-100 text-gray-600"}`}>
      {status}
    </span>
  );
}

function AlertBadge({ dias, valor }) {
  if (dias >= 60) return <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-bold bg-red-600 text-white">🔴 {dias}d CRÍTICO</span>;
  if (dias >= 50) return <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-bold bg-red-100 text-red-700">🔴 {dias}d</span>;
  if (dias >= 25) return <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-bold bg-yellow-100 text-yellow-800">🟡 {dias}d</span>;
  return <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold bg-green-100 text-green-700">🟢 {dias}d</span>;
}

// ============================================================
// MODAL DETALHE DA NOTA
// ============================================================
function ModalNota({ nota, onClose, onSave, usuarioAtual }) {
  const [form, setForm] = useState({ ...nota });
  const [activeTab, setActiveTab] = useState("dados");

  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  const salvar = () => {
    const agora = new Date().toLocaleString("pt-BR");
    const acoes = [];
    if (form.status !== nota.status) acoes.push(`Status alterado: "${nota.status}" → "${form.status}"`);
    if (form.centroCusto !== nota.centroCusto) acoes.push(`Centro de Custo: "${form.centroCusto}"`);
    if (form.finalidade !== nota.finalidade) acoes.push(`Finalidade: "${form.finalidade}"`);
    if (form.responsavel !== nota.responsavel) acoes.push(`Responsável: "${form.responsavel}"`);
    if (form.noRM !== nota.noRM) acoes.push(`No RM: ${form.noRM ? "Sim" : "Não"}`);

    const novoHistorico = [...(nota.historico || [])];
    acoes.forEach(a => novoHistorico.push({ acao: a, usuario: usuarioAtual.nome, data: agora }));
    if (acoes.length === 0) novoHistorico.push({ acao: "Nota revisada sem alterações", usuario: usuarioAtual.nome, data: agora });

    onSave({ ...form, historico: novoHistorico });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background: "rgba(0,0,0,0.5)" }}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-3xl max-h-screen overflow-y-auto m-4">
        {/* Header */}
        <div className="flex items-center justify-between p-5 border-b" style={{ borderColor: "#f0f0f0" }}>
          <div>
            <h2 className="font-bold text-lg text-gray-800">{nota.razaoSocial}</h2>
            <p className="text-sm text-gray-500">NF-e nº {nota.numNota} — CFOP {nota.cfop} — {EMPRESAS.find(e => e.id === nota.empresa)?.nome}</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 text-2xl font-light">×</button>
        </div>

        {/* Próximo Passo */}
        <div className="mx-5 mt-4 p-3 rounded-xl text-sm font-medium" style={{ background: "#f0f8f8", border: "1px solid #ffd6b8", color: "#b84a00" }}>
          {getProximoPasso(form)}
        </div>

        {/* Alertas */}
        {getAlertas(form).filter(a => a.tipo !== "ok").map((al, i) => (
          <div key={i} className="mx-5 mt-2 p-3 rounded-xl text-sm font-medium" style={{ background: al.tipo === "danger" ? "#fff0f0" : "#fffbeb", border: `1px solid ${al.tipo === "danger" ? "#ffc7c7" : "#fde68a"}`, color: al.tipo === "danger" ? "#c0392b" : "#92400e" }}>
            {al.msg}
          </div>
        ))}

        {/* Tabs */}
        <div className="flex gap-1 px-5 mt-4 border-b" style={{ borderColor: "#f0f0f0" }}>
          {["dados", "financeiro", "historico"].map(t => (
            <button key={t} onClick={() => setActiveTab(t)}
              className={`px-4 py-2 text-sm font-semibold capitalize rounded-t-lg transition-all ${activeTab === t ? "text-white" : "text-gray-500 hover:text-gray-700"}`}
              style={activeTab === t ? { background: "#E8450A" } : {}}>
              {t === "dados" ? "Dados" : t === "financeiro" ? "Financeiro" : "Histórico"}
            </button>
          ))}
        </div>

        <div className="p-5">
          {activeTab === "dados" && (
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="text-xs font-semibold text-gray-500 uppercase">CNPJ Fornecedor</label>
                <p className="text-sm text-gray-800 mt-1">{form.fornecedor}</p>
              </div>
              <div>
                <label className="text-xs font-semibold text-gray-500 uppercase">Chave NF-e</label>
                <p className="text-xs text-gray-600 mt-1 break-all">{form.chave}</p>
              </div>
              <div>
                <label className="text-xs font-semibold text-gray-500 uppercase">Dt. Emissão</label>
                <p className="text-sm text-gray-800 mt-1">{form.dtEmissao}</p>
              </div>
              <div>
                <label className="text-xs font-semibold text-gray-500 uppercase">Valor</label>
                <p className="text-sm font-bold mt-1" style={{ color: form.valor > 25000 ? "#c0392b" : "#2d6a4f" }}>{fmtMoeda(form.valor)}</p>
              </div>
              <div>
                <label className="text-xs font-semibold text-gray-500 uppercase">Status</label>
                <select value={form.status} onChange={e => set("status", e.target.value)}
                  className="mt-1 w-full border rounded-lg p-2 text-sm" style={{ borderColor: "#e5e7eb" }}>
                  {STATUS_LIST.map(s => <option key={s}>{s}</option>)}
                </select>
              </div>
              <div>
                <label className="text-xs font-semibold text-gray-500 uppercase">Nota no RM?</label>
                <select value={form.noRM === null ? "" : form.noRM ? "sim" : "nao"} onChange={e => set("noRM", e.target.value === "" ? null : e.target.value === "sim")}
                  className="mt-1 w-full border rounded-lg p-2 text-sm" style={{ borderColor: "#e5e7eb" }}>
                  <option value="">Não verificado</option>
                  <option value="sim">✅ Sim</option>
                  <option value="nao">❌ Não</option>
                </select>
              </div>
              <div>
                <label className="text-xs font-semibold text-gray-500 uppercase">Código do CC</label>
                <input value={form.codigoCusto || ""} onChange={e => set("codigoCusto", e.target.value)}
                  placeholder="Ex: 1050"
                  className="mt-1 w-full border rounded-lg p-2 text-sm" style={{ borderColor: "#e5e7eb" }} />
              </div>
              <div>
                <label className="text-xs font-semibold text-gray-500 uppercase">Centro de Custo</label>
                <input value={form.centroCusto} onChange={e => set("centroCusto", e.target.value)}
                  placeholder="Ex: Obra MLA - Contrato 001"
                  className="mt-1 w-full border rounded-lg p-2 text-sm" style={{ borderColor: "#e5e7eb" }} />
              </div>
              <div>
                <label className="text-xs font-semibold text-gray-500 uppercase">Finalidade</label>
                <select value={form.finalidade} onChange={e => set("finalidade", e.target.value)}
                  className="mt-1 w-full border rounded-lg p-2 text-sm" style={{ borderColor: "#e5e7eb" }}>
                  <option value="">Selecione...</option>
                  {FINALIDADES.map(f => <option key={f}>{f}</option>)}
                </select>
              </div>
              <div className="col-span-2">
                <label className="text-xs font-semibold text-gray-500 uppercase">Responsável / Comprador</label>
                <input value={form.responsavel} onChange={e => set("responsavel", e.target.value)}
                  placeholder="Nome do responsável"
                  className="mt-1 w-full border rounded-lg p-2 text-sm" style={{ borderColor: "#e5e7eb" }} />
              </div>
              <div>
                <label className="text-xs font-semibold text-gray-500 uppercase">Dt. Reanálise</label>
                <input type="text" value={form.dtReanalise} onChange={e => set("dtReanalise", e.target.value)}
                  placeholder="DD/MM/AAAA"
                  className="mt-1 w-full border rounded-lg p-2 text-sm" style={{ borderColor: "#e5e7eb" }} />
              </div>
              <div>
                <label className="text-xs font-semibold text-gray-500 uppercase">Dt. Postergação</label>
                <input type="text" value={form.dtPostergacao} onChange={e => set("dtPostergacao", e.target.value)}
                  placeholder="DD/MM/AAAA"
                  className="mt-1 w-full border rounded-lg p-2 text-sm" style={{ borderColor: "#e5e7eb" }} />
              </div>
              {form.status === "Selada" && (
                <div className="col-span-2">
                  <label className="text-xs font-semibold uppercase" style={{ color: "#2d6a4f" }}>📌 Data de Selagem <span className="text-red-400">*</span></label>
                  <input type="text" value={form.dtSelada || ""} onChange={e => { let v = e.target.value.replace(/\D/g,""); if(v.length>=3) v=v.slice(0,2)+"/"+v.slice(2); if(v.length>=6) v=v.slice(0,5)+"/"+v.slice(5); set("dtSelada", v.slice(0,10)); }}
                    placeholder="DD/MM/AAAA — data em que foi selada no posto fiscal"
                    className="mt-1 w-full border-2 rounded-lg p-2 text-sm font-medium"
                    style={{ borderColor: "#2d6a4f", background: "#f0fdf4" }} />
                  <p className="text-xs mt-1" style={{ color: "#2d6a4f" }}>Nota apresentada no posto fiscal e selada com sucesso.</p>
                </div>
              )}
              <div className="col-span-2">
                <label className="text-xs font-semibold text-gray-500 uppercase">Observações</label>
                <textarea value={form.obs} onChange={e => set("obs", e.target.value)} rows={3}
                  className="mt-1 w-full border rounded-lg p-2 text-sm" style={{ borderColor: "#e5e7eb" }} />
              </div>
            </div>
          )}

          {activeTab === "financeiro" && (
            <AbaFinanceiro form={form} set={set} fmtMoeda={fmtMoeda} />
          )}

          {activeTab === "historico" && (
            <div className="space-y-2">
              {(form.historico || []).slice().reverse().map((h, i) => (
                <div key={i} className="flex gap-3 p-3 rounded-lg" style={{ background: "#f8f9fa" }}>
                  <div className="w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold text-white flex-shrink-0" style={{ background: "#E8450A" }}>
                    {h.usuario?.charAt(0) || "S"}
                  </div>
                  <div>
                    <p className="text-sm text-gray-800">{h.acao}</p>
                    <p className="text-xs text-gray-400 mt-0.5">{h.usuario} — {h.data}</p>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        <div className="flex justify-end gap-3 p-5 border-t" style={{ borderColor: "#f0f0f0" }}>
          <button onClick={onClose} className="px-4 py-2 rounded-lg border text-sm text-gray-600 hover:bg-gray-50" style={{ borderColor: "#e5e7eb" }}>Cancelar</button>
          <button onClick={salvar} className="px-6 py-2 rounded-lg text-sm font-semibold text-white" style={{ background: "#E8450A" }}>Salvar Alterações</button>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// DRAWER DE ALERTAS
// ============================================================
function DrawerAlertas({ notas, filtro, onClose, onVerNota }) {
  const ativas = notas.filter(n => !["Desembaraçada", "Selada", "Recusada"].includes(n.status));
  let lista = [];
  let titulo = "";

  const agora2 = new Date();
  const mesAtual2 = agora2.getMonth();
  const anoAtual2 = agora2.getFullYear();
  const nomeMes2 = agora2.toLocaleString("pt-BR", { month: "long" });
  function noMesAtual2(dtStr) {
    if (!dtStr || dtStr === "-") return false;
    try {
      let dt;
      if (dtStr.includes("/")) {
        const p = dtStr.split("/");
        if (p.length !== 3) return false;
        dt = new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
      } else {
        dt = new Date(dtStr); // ISO format
      }
      if (isNaN(dt.getTime())) return false;
      return dt.getMonth() === mesAtual2 && dt.getFullYear() === anoAtual2;
    } catch { return false; }
  }

  if (filtro === "critico") {
    lista = ativas.filter(n => n.qtdeDias >= 60).sort((a, b) => b.qtdeDias - a.qtdeDias);
    titulo = "🔴 Notas Críticas — 60+ dias";
  } else if (filtro === "atencao") {
    lista = ativas.filter(n => n.qtdeDias >= 25 && n.qtdeDias < 60).sort((a, b) => b.qtdeDias - a.qtdeDias);
    titulo = "🟡 Atenção — 25 a 59 dias";
  } else if (filtro === "acima25k") {
    lista = ativas.filter(n => n.valor > 25000).sort((a, b) => b.valor - a.valor);
    titulo = "⚠️ Notas acima de R$ 25.000";
  } else if (filtro === "pagamento") {
    lista = ativas.filter(n => n.status === "Aguardando Pagamento");
    titulo = "💰 Aguardando Pagamento";
  } else if (filtro === "concluidoMes") {
    lista = notas.filter(n => (n.status === "Desembaraçada" || n.status === "Selada") && noMesAtual2(n.dtImportacao))
      .sort((a, b) => a.razaoSocial.localeCompare(b.razaoSocial));
    titulo = "✅ Concluídas em " + nomeMes2;
  } else if (filtro === "custoMes") {
    lista = notas.filter(n => noMesAtual2(n.dtImportacao) && (n.taxaReanalise||0)+(n.taxaDesembaraco||0)+(n.icmsAntecipado||0)+(n.multa10pct||0)+(n.multa||0)+(n.juros||0) > 0)
      .sort((a, b) => ((b.taxaReanalise||0)+(b.taxaDesembaraco||0)+(b.icmsAntecipado||0)+(b.multa10pct||0)+(b.multa||0)+(b.juros||0)) - ((a.taxaReanalise||0)+(a.taxaDesembaraco||0)+(a.icmsAntecipado||0)+(a.multa10pct||0)+(a.multa||0)+(a.juros||0)));
    titulo = "📊 Custos em " + nomeMes2;
  } else if (filtro === "comCusto") {
    lista = ativas.filter(n => (n.taxaReanalise||0)+(n.taxaDesembaraco||0)+(n.icmsAntecipado||0)+(n.multa||0)+(n.juros||0) > 0)
      .sort((a, b) => ((b.taxaReanalise||0)+(b.taxaDesembaraco||0)+(b.icmsAntecipado||0)+(b.multa||0)+(b.juros||0)) - ((a.taxaReanalise||0)+(a.taxaDesembaraco||0)+(a.icmsAntecipado||0)+(a.multa||0)+(a.juros||0)));
    titulo = "📊 Custos Registrados — Notas Ativas";
  } else if (filtro?.startsWith("empresa_")) {
    const empId = filtro.replace("empresa_", "");
    lista = ativas.filter(n => n.empresa === empId).sort((a, b) => b.qtdeDias - a.qtdeDias);
    titulo = `🏢 ${EMPRESAS.find(e => e.id === empId)?.nome}`;
  }

  return (
    <div className="fixed inset-0 z-50 flex justify-end" style={{ background: "rgba(0,0,0,0.4)" }} onClick={onClose}>
      <div className="bg-white h-full w-full max-w-md flex flex-col shadow-2xl" onClick={e => e.stopPropagation()}>
        <div className="flex items-center justify-between p-5 border-b" style={{ borderColor: "#f0f0f0" }}>
          <div>
            <h2 className="font-bold text-gray-800">{titulo}</h2>
            <p className="text-xs text-gray-400 mt-0.5">{lista.length} nota(s)</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 text-2xl font-light leading-none">×</button>
        </div>
        <div className="flex-1 overflow-y-auto p-4 space-y-3">
          {lista.length === 0 && (
            <div className="text-center py-12 text-gray-400">
              <p className="text-3xl mb-2">✅</p>
              <p className="text-sm">Nenhuma nota nesta categoria</p>
            </div>
          )}
          {lista.map(n => (
            <div key={n.id} className="p-4 rounded-xl border cursor-pointer hover:shadow-md transition-all"
              style={{ borderColor: n.qtdeDias >= 60 ? "#ffc7c7" : n.qtdeDias >= 25 ? "#fde68a" : "#e5e7eb", background: n.qtdeDias >= 60 ? "#fff8f8" : n.qtdeDias >= 25 ? "#fffdf0" : "#fff" }}
              onClick={() => { onVerNota(n); onClose(); }}>
              <div className="flex items-center justify-between mb-2">
                <AlertBadge dias={n.qtdeDias} />
                <span className="font-bold text-sm" style={{ color: n.valor > 25000 ? "#c0392b" : "#374151" }}>{fmtMoeda(n.valor)}</span>
              </div>
              <p className="font-semibold text-sm text-gray-800 truncate">{n.razaoSocial}</p>
              <p className="text-xs text-gray-400 mt-0.5">NF {n.numNota} • {EMPRESAS.find(e => e.id === n.empresa)?.nome}</p>
              <p className="text-xs mt-2 font-medium" style={{ color: "#E8450A" }}>{getProximoPasso(n)}</p>
              <div className="mt-2"><Badge status={n.status} /></div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ============================================================
// TELA: DASHBOARD
// ============================================================
function Dashboard({ notas, onVerNota, onIrParaPainel }) {
  const [drawerFiltro, setDrawerFiltro] = useState(null);

  // Mês atual
  const agora = new Date();
  const mesAtual = agora.getMonth();
  const anoAtual = agora.getFullYear();

  function noMesAtual(dtStr) {
    if (!dtStr || dtStr === "-") return false;
    try {
      let dt;
      if (dtStr.includes("/")) {
        const p = dtStr.split("/");
        if (p.length !== 3) return false;
        dt = new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
      } else {
        dt = new Date(dtStr); // ISO format
      }
      if (isNaN(dt.getTime())) return false;
      return dt.getMonth() === mesAtual && dt.getFullYear() === anoAtual;
    } catch { return false; }
  }

  const nomeMes = agora.toLocaleString("pt-BR", { month: "long" });

  const statusConcluido = ["Desembaraçada", "Selada", "Recusada"];
  const ativas = notas.filter(n => !statusConcluido.includes(n.status));
  const criticas = ativas.filter(n => n.qtdeDias >= 60);
  const atencao = ativas.filter(n => n.qtdeDias >= 25 && n.qtdeDias < 60);
  const acima25k = ativas.filter(n => n.valor > 25000);
  const aguardPag = ativas.filter(n => n.status === "Aguardando Pagamento");

  // Desembaraçadas + Seladas no mês atual (pela data de importação)
  const desembaracadasMes = notas.filter(n =>
    (n.status === "Desembaraçada" || n.status === "Selada") && noMesAtual(n.dtImportacao)
  );

  // Custos do mês atual — todas as notas importadas no mês, independente do status
  const notasMes = notas.filter(n => noMesAtual(n.dtImportacao));
  const totalCustosMes = notasMes.reduce((s, n) =>
    s + (n.taxaReanalise||0) + (n.taxaDesembaraco||0) + (n.icmsAntecipado||0) + (n.multa10pct||0) + (n.multa||0) + (n.juros||0), 0);

  const porEmpresa = EMPRESAS.map(e => ({
    ...e,
    total: ativas.filter(n => n.empresa === e.id).length,
    valor: ativas.filter(n => n.empresa === e.id).reduce((s, n) => s + n.valor, 0),
    criticas: ativas.filter(n => n.empresa === e.id && n.qtdeDias >= 60).length,
    atencao: ativas.filter(n => n.empresa === e.id && n.qtdeDias >= 25 && n.qtdeDias < 60).length,
  }));

  const cards = [
    { label: "Notas Críticas", valor: criticas.length, sub: "60+ dias • clique para ver", cor: "#c0392b", bg: "#fff0f0", icon: "🔴", filtro: "critico", clicavel: true },
    { label: "Atenção", valor: atencao.length, sub: "25–59 dias • clique para ver", cor: "#b7791f", bg: "#fffbeb", icon: "🟡", filtro: "atencao", clicavel: true },
    { label: "Acima de R$ 25k", valor: acima25k.length, sub: "Risco multa 10% • clique para ver", cor: "#7e3af2", bg: "#f5f3ff", icon: "⚠️", filtro: "acima25k", clicavel: true },
    { label: "Aguard. Pagamento", valor: aguardPag.length, sub: "Financeiro pendente • clique para ver", cor: "#1a56db", bg: "#eff6ff", icon: "💰", filtro: "pagamento", clicavel: true },
    { label: "Concluídas em " + nomeMes, valor: desembaracadasMes.length, sub: "Desembaraçadas + Seladas no mês", cor: "#2d6a4f", bg: "#f0fdf4", icon: "✅", filtro: "concluidoMes", clicavel: desembaracadasMes.length > 0 },
    { label: "Custos em " + nomeMes, valor: fmtMoeda(totalCustosMes), sub: "Notas importadas no mês • clique para ver", cor: "#E8450A", bg: "#f0f8f8", icon: "📊", filtro: "custoMes", clicavel: true },
  ];

  return (
    <div className="space-y-6">
      {/* Cards resumo — todos clicáveis onde aplicável */}
      <div className="grid grid-cols-2 gap-4" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))" }}>
        {cards.map((c, i) => (
          <div key={i}
            onClick={() => c.clicavel && setDrawerFiltro(c.filtro)}
            className={`rounded-2xl p-4 transition-all ${c.clicavel ? "cursor-pointer hover:shadow-md hover:-translate-y-0.5" : ""}`}
            style={{ background: c.bg, border: `1px solid ${c.cor}33` }}>
            <div className="flex items-center justify-between">
              <span className="text-2xl">{c.icon}</span>
              <div className="flex items-center gap-2">
                <span className="text-2xl font-black" style={{ color: c.cor }}>{c.valor}</span>
                {c.clicavel && <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: c.cor + "22", color: c.cor }}>ver →</span>}
              </div>
            </div>
            <p className="font-semibold text-gray-700 mt-2 text-sm">{c.label}</p>
            <p className="text-xs text-gray-400">{c.sub}</p>
          </div>
        ))}
      </div>

      {/* Por empresa — cada linha clicável */}
      <div className="rounded-2xl border p-5" style={{ borderColor: "#f0f0f0" }}>
        <h3 className="font-bold text-gray-700 mb-1 text-sm uppercase tracking-wide">Notas Ativas por Empresa</h3>
        <p className="text-xs text-gray-400 mb-4">Clique em uma empresa para ver as notas filtradas</p>
        <div className="space-y-3">
          {porEmpresa.map(e => (
            <div key={e.id}
              onClick={() => setDrawerFiltro(`empresa_${e.id}`)}
              className="flex items-center justify-between p-4 rounded-xl cursor-pointer hover:shadow-md hover:-translate-y-0.5 transition-all"
              style={{ background: "#f8f9fa", border: "1px solid #eee" }}>
              <div className="flex-1">
                <p className="font-semibold text-sm text-gray-800">{e.nome}</p>
                <div className="flex items-center gap-3 mt-1">
                  <span className="text-xs text-gray-400">{e.total} nota(s) ativa(s)</span>
                  {e.criticas > 0 && <span className="text-xs font-bold px-2 py-0.5 rounded-full bg-red-100 text-red-600">{e.criticas} crítica(s)</span>}
                  {e.atencao > 0 && <span className="text-xs font-bold px-2 py-0.5 rounded-full bg-yellow-100 text-yellow-700">{e.atencao} atenção</span>}
                </div>
              </div>
              <div className="text-right ml-3">
                <p className="font-bold text-sm" style={{ color: "#E8450A" }}>{fmtMoeda(e.valor)}</p>
                <p className="text-xs text-gray-400">em exposição</p>
                <p className="text-xs font-semibold mt-1" style={{ color: "#E8450A" }}>ver notas →</p>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Botão flutuante de alertas */}
      {(criticas.length > 0 || atencao.length > 0) && (
        <div className="fixed bottom-6 right-6 z-40 flex flex-col gap-2 items-end">
          {criticas.length > 0 && (
            <button onClick={() => setDrawerFiltro("critico")}
              className="flex items-center gap-2 px-4 py-3 rounded-full text-white text-sm font-bold shadow-lg hover:shadow-xl transition-all hover:-translate-y-0.5"
              style={{ background: "#c0392b" }}>
              🔴 {criticas.length} nota(s) crítica(s)
            </button>
          )}
          {atencao.length > 0 && (
            <button onClick={() => setDrawerFiltro("atencao")}
              className="flex items-center gap-2 px-4 py-3 rounded-full text-sm font-bold shadow-lg hover:shadow-xl transition-all hover:-translate-y-0.5"
              style={{ background: "#b7791f", color: "white" }}>
              🟡 {atencao.length} nota(s) em atenção
            </button>
          )}
        </div>
      )}

      {/* Drawer lateral */}
      {drawerFiltro && (
        <DrawerAlertas
          notas={notas}
          filtro={drawerFiltro}
          onClose={() => setDrawerFiltro(null)}
          onVerNota={onVerNota}
        />
      )}
    </div>
  );
}

// ============================================================
// TELA: PAINEL DE NOTAS
// ============================================================
function PainelNotas({ notas, onVerNota, onImportar, ultimaImportacao, empresas }) {
  const [filtros, setFiltros] = useState({
    empresa: "", status: "", busca: "", dias: "", valor: "",
    tipoData: "emissao", dtDe: "", dtAte: ""
  });
  const [filtrosAplicados, setFiltrosAplicados] = useState({
    empresa: "", status: "", busca: "", dias: "", valor: "",
    tipoData: "emissao", dtDe: "", dtAte: ""
  });
  const [mostrarImport, setMostrarImport] = useState(false);
  const setF = (k, v) => setFiltros(f => ({ ...f, [k]: v }));
  const aplicarBusca = () => setFiltrosAplicados({ ...filtros });
  const limparTudo = () => {
    const vazio = { empresa: "", status: "", busca: "", dias: "", valor: "", tipoData: "emissao", dtDe: "", dtAte: "" };
    setFiltros(vazio);
    setFiltrosAplicados(vazio);
  };

  function parseDateBR2(str) {
    if (!str || str.length !== 10) return null;
    const [d, m, y] = str.split("/");
    if (!d || !m || !y) return null;
    return new Date(`${y}-${m}-${d}`);
  }

  const filtradas = notas.filter(n => {
    const f = filtrosAplicados;
    if (f.empresa && n.empresa !== f.empresa) return false;
    if (f.status && n.status !== f.status) return false;
    if (f.dias === "critico" && n.qtdeDias < 60) return false;
    if (f.dias === "atencao" && (n.qtdeDias < 25 || n.qtdeDias >= 60)) return false;
    if (f.dias === "ok" && n.qtdeDias >= 25) return false;
    if (f.valor === "acima25k" && n.valor <= 25000) return false;
    if (f.valor === "ate25k" && n.valor > 25000) return false;
    if (f.busca) {
      const b = f.busca.toLowerCase();
      if (!n.razaoSocial.toLowerCase().includes(b) && !n.numNota.includes(b) && !n.fornecedor.includes(b) && !n.chave.includes(b)) return false;
    }
    if (f.dtDe || f.dtAte) {
      const campoData = f.tipoData === "emissao" ? n.dtEmissao : n.dtImportacao;
      const dataRef = parseDateBR2(campoData);
      const de = parseDateBR2(f.dtDe);
      const ate = parseDateBR2(f.dtAte);
      if (!dataRef) return false;
      if (de && dataRef < de) return false;
      if (ate && dataRef > ate) return false;
    }
    return true;
  });

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => onImportar(ev.target.result, file.name);
    reader.readAsText(file, "ascii");
    e.target.value = "";
    setMostrarImport(false);
  };

  const temFiltroAtivo = filtrosAplicados.empresa || filtrosAplicados.status || filtrosAplicados.busca || filtrosAplicados.dias || filtrosAplicados.valor || filtrosAplicados.dtDe || filtrosAplicados.dtAte;

  return (
    <div className="space-y-4">

      {/* Banner de importação destacado */}
      <div className="rounded-2xl p-4 flex items-center justify-between" style={{ background: "#f0f8f8", border: "2px dashed #E8450A" }}>
        <div>
          <p className="font-bold text-sm" style={{ color: "#1a4a4a" }}>📥 Importar arquivo do DTE</p>
          <p className="text-xs text-gray-500 mt-0.5">
            Faça upload do CSV/Excel exportado do DTE semanalmente.
            {ultimaImportacao && <span className="ml-2 font-medium text-gray-600">Última importação: {ultimaImportacao}</span>}
          </p>
        </div>
        <div className="flex items-center gap-2 ml-4">
          <button onClick={() => exportarExcel(filtradas, empresas || EMPRESAS, "painel")}
            className="px-4 py-2.5 rounded-xl text-sm font-bold border whitespace-nowrap" style={{ borderColor: "#E8450A", color: "#E8450A", background: "white" }}>
            📊 Exportar Excel
          </button>
          <label className="px-5 py-2.5 rounded-xl text-sm font-bold text-white cursor-pointer whitespace-nowrap hover:opacity-90 transition-opacity" style={{ background: "#E8450A" }}>
            ⬆ Selecionar Arquivo
            <input type="file" accept=".csv,.xlsx" className="hidden" onChange={handleFile} />
          </label>
        </div>
      </div>

      {/* Filtros */}
      <div className="rounded-2xl border p-4 space-y-3" style={{ borderColor: "#f0f0f0", background: "#fafafa" }}>
        <div className="flex flex-wrap gap-2 items-end">
          <div className="flex flex-col gap-1" style={{ minWidth: 220 }}>
            <label className="text-xs font-semibold text-gray-500">Busca</label>
            <input placeholder="Fornecedor, NF ou chave..." value={filtros.busca} onChange={e => setF("busca", e.target.value)}
              onKeyDown={e => e.key === "Enter" && aplicarBusca()}
              className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: "#e5e7eb" }} />
          </div>
          <div className="flex flex-col gap-1">
            <label className="text-xs font-semibold text-gray-500">Empresa</label>
            <select value={filtros.empresa} onChange={e => setF("empresa", e.target.value)} className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: "#e5e7eb" }}>
              <option value="">Todas</option>
              {EMPRESAS.map(e => <option key={e.id} value={e.id}>{e.nome}</option>)}
            </select>
          </div>
          <div className="flex flex-col gap-1">
            <label className="text-xs font-semibold text-gray-500">Status</label>
            <select value={filtros.status} onChange={e => setF("status", e.target.value)} className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: "#e5e7eb" }}>
              <option value="">Todos</option>
              {STATUS_LIST.map(s => <option key={s}>{s}</option>)}
            </select>
          </div>
          <div className="flex flex-col gap-1">
            <label className="text-xs font-semibold text-gray-500">Prazo</label>
            <select value={filtros.dias} onChange={e => setF("dias", e.target.value)} className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: "#e5e7eb" }}>
              <option value="">Todos</option>
              <option value="critico">🔴 Crítico 60+d</option>
              <option value="atencao">🟡 Atenção 25-59d</option>
              <option value="ok">🟢 OK 0-24d</option>
            </select>
          </div>
          <div className="flex flex-col gap-1">
            <label className="text-xs font-semibold text-gray-500">Valor</label>
            <select value={filtros.valor} onChange={e => setF("valor", e.target.value)} className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: "#e5e7eb" }}>
              <option value="">Todos</option>
              <option value="acima25k">Acima R$ 25k</option>
              <option value="ate25k">Até R$ 25k</option>
            </select>
          </div>
          <div className="flex flex-col gap-1">
            <label className="text-xs font-semibold text-gray-500">Data</label>
            <select value={filtros.tipoData} onChange={e => setF("tipoData", e.target.value)} className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: "#e5e7eb" }}>
              <option value="emissao">Emissão</option>
              <option value="importacao">Importação</option>
            </select>
          </div>
          <div className="flex flex-col gap-1">
            <label className="text-xs font-semibold text-gray-500">De</label>
            <input type="text" placeholder="DD/MM/AAAA" value={filtros.dtDe} maxLength={10}
              onChange={e => { let v=e.target.value.replace(/\D/g,""); if(v.length>=3)v=v.slice(0,2)+"/"+v.slice(2); if(v.length>=6)v=v.slice(0,5)+"/"+v.slice(5); setF("dtDe",v.slice(0,10)); }}
              className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: "#e5e7eb", width: 110 }} />
          </div>
          <div className="flex flex-col gap-1">
            <label className="text-xs font-semibold text-gray-500">Até</label>
            <input type="text" placeholder="DD/MM/AAAA" value={filtros.dtAte} maxLength={10}
              onChange={e => { let v=e.target.value.replace(/\D/g,""); if(v.length>=3)v=v.slice(0,2)+"/"+v.slice(2); if(v.length>=6)v=v.slice(0,5)+"/"+v.slice(5); setF("dtAte",v.slice(0,10)); }}
              className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: "#e5e7eb", width: 110 }} />
          </div>
          <button onClick={aplicarBusca}
            className="px-5 py-2 rounded-lg text-sm font-bold text-white flex items-center gap-2"
            style={{ background: "#1a4a4a" }}>
            🔍 Buscar
          </button>
          {(filtros.empresa||filtros.status||filtros.busca||filtros.dias||filtros.valor||filtros.dtDe||filtros.dtAte) && (
            <button onClick={limparTudo} className="px-4 py-2 rounded-lg text-sm border font-medium text-gray-500" style={{ borderColor: "#e5e7eb" }}>
              Limpar
            </button>
          )}
        </div>
      </div>

      <p className="text-xs text-gray-400 px-1">
        {filtradas.length} nota(s) encontrada(s)
        {temFiltroAtivo && <span className="ml-2 font-semibold" style={{ color: "#E8450A" }}>• filtros ativos</span>}
      </p>

      {/* Cards de notas */}
      <div className="space-y-2">
        {filtradas.length === 0 && (
          <div className="text-center py-12 text-gray-400 rounded-2xl border" style={{ borderColor: "#f0f0f0" }}>
            <p className="text-4xl mb-2">📋</p>
            <p>Nenhuma nota encontrada com os filtros selecionados</p>
          </div>
        )}
        {filtradas.map((n) => {
          const emp = EMPRESAS.find(e => e.id === n.empresa);
          const empNome = emp?.nome?.split(" - ").pop() || n.empresa;
          const diasCor = n.qtdeDias >= 60 ? { bg: "#fee2e2", txt: "#991b1b" } : n.qtdeDias >= 25 ? { bg: "#fef9c3", txt: "#854d0e" } : { bg: "#dcfce7", txt: "#166534" };
          return (
            <div key={n.id} className="rounded-xl border p-3 flex items-center gap-3 hover:border-teal-300 transition-colors" style={{ borderColor: "#f0f0f0", background: "#fff" }}>
              <span className="text-xs font-medium px-2 py-1 rounded-full flex-shrink-0" style={{ background: "#edf5f5", color: "#1a4a4a" }}>{empNome}</span>
              <div className="flex-1 min-w-0">
                <p className="font-semibold text-sm text-gray-800 truncate">{n.razaoSocial}</p>
                <p className="text-xs text-gray-400 mt-0.5">NF {n.numNota} · CFOP {n.cfop} · {n.dtEmissao}</p>
              </div>
              <span className="text-sm font-bold flex-shrink-0" style={{ color: n.valor > 25000 ? "#c0392b" : "#374151" }}>{fmtMoeda(n.valor)}</span>
              <span className="text-xs font-bold px-2 py-1 rounded-md flex-shrink-0" style={{ background: diasCor.bg, color: diasCor.txt }}>{n.qtdeDias}d</span>
              <Badge status={n.status} />
              {n.centroCusto && <span className="text-xs text-gray-400 hidden md:block flex-shrink-0 max-w-24 truncate">{n.centroCusto}</span>}
              <button onClick={() => onVerNota(n)} className="px-3 py-1.5 rounded-lg text-xs font-bold text-white flex-shrink-0" style={{ background: "#E8450A" }}>
                ✏ Editar
              </button>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ============================================================
// SVG TORRES DE TRANSMISSÃO (background decorativo)
// ============================================================
function TorresSVG({ opacity = 0.07, corTorre = "#E8450A" }) {
  return (
    <svg viewBox="0 0 1200 400" xmlns="http://www.w3.org/2000/svg"
      style={{ position: "absolute", bottom: 0, left: 0, width: "100%", height: "100%", opacity }}
      preserveAspectRatio="xMidYMax meet">
      <defs>
        <linearGradient id="torreGrad" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor={corTorre} stopOpacity="1"/>
          <stop offset="100%" stopColor={corTorre} stopOpacity="0.6"/>
        </linearGradient>
      </defs>

      {/* TORRE 1 — esquerda pequena */}
      <g transform="translate(30,80)">
        <line x1="35" y1="0" x2="18" y2="320" stroke="url(#torreGrad)" strokeWidth="2.5"/>
        <line x1="35" y1="0" x2="52" y2="320" stroke="url(#torreGrad)" strokeWidth="2.5"/>
        <line x1="18" y1="320" x2="52" y2="320" stroke={corTorre} strokeWidth="2.5"/>
        {/* diagonais internas */}
        <line x1="18" y1="320" x2="35" y2="200" stroke={corTorre} strokeWidth="1.2"/>
        <line x1="52" y1="320" x2="35" y2="200" stroke={corTorre} strokeWidth="1.2"/>
        <line x1="18" y1="200" x2="35" y2="120" stroke={corTorre} strokeWidth="1.2"/>
        <line x1="52" y1="200" x2="35" y2="120" stroke={corTorre} strokeWidth="1.2"/>
        <line x1="22" y1="120" x2="35" y2="60" stroke={corTorre} strokeWidth="1.2"/>
        <line x1="48" y1="120" x2="35" y2="60" stroke={corTorre} strokeWidth="1.2"/>
        {/* travessas horizontais */}
        <line x1="10" y1="200" x2="60" y2="200" stroke={corTorre} strokeWidth="1.8"/>
        <line x1="12" y1="120" x2="58" y2="120" stroke={corTorre} strokeWidth="1.8"/>
        <line x1="15" y1="60" x2="55" y2="60" stroke={corTorre} strokeWidth="2"/>
        {/* braços do topo */}
        <line x1="-15" y1="28" x2="85" y2="28" stroke={corTorre} strokeWidth="2.2"/>
        <line x1="-8" y1="10" x2="78" y2="10" stroke={corTorre} strokeWidth="2.2"/>
        <line x1="35" y1="0" x2="35" y2="10" stroke={corTorre} strokeWidth="2"/>
        {/* isoladores */}
        <circle cx="-15" cy="28" r="3" fill={corTorre}/>
        <circle cx="85" cy="28" r="3" fill={corTorre}/>
        <circle cx="-8" cy="10" r="3" fill={corTorre}/>
        <circle cx="78" cy="10" r="3" fill={corTorre}/>
        {/* base */}
        <line x1="5" y1="310" x2="18" y2="320" stroke={corTorre} strokeWidth="2"/>
        <line x1="65" y1="310" x2="52" y2="320" stroke={corTorre} strokeWidth="2"/>
      </g>

      {/* TORRE 2 — média */}
      <g transform="translate(280,40)">
        <line x1="45" y1="0" x2="22" y2="360" stroke="url(#torreGrad)" strokeWidth="3"/>
        <line x1="45" y1="0" x2="68" y2="360" stroke="url(#torreGrad)" strokeWidth="3"/>
        <line x1="22" y1="360" x2="68" y2="360" stroke={corTorre} strokeWidth="3"/>
        <line x1="22" y1="360" x2="45" y2="240" stroke={corTorre} strokeWidth="1.4"/>
        <line x1="68" y1="360" x2="45" y2="240" stroke={corTorre} strokeWidth="1.4"/>
        <line x1="22" y1="240" x2="45" y2="150" stroke={corTorre} strokeWidth="1.4"/>
        <line x1="68" y1="240" x2="45" y2="150" stroke={corTorre} strokeWidth="1.4"/>
        <line x1="26" y1="150" x2="45" y2="75" stroke={corTorre} strokeWidth="1.4"/>
        <line x1="64" y1="150" x2="45" y2="75" stroke={corTorre} strokeWidth="1.4"/>
        <line x1="8" y1="240" x2="82" y2="240" stroke={corTorre} strokeWidth="2"/>
        <line x1="10" y1="150" x2="80" y2="150" stroke={corTorre} strokeWidth="2"/>
        <line x1="13" y1="75" x2="77" y2="75" stroke={corTorre} strokeWidth="2.2"/>
        <line x1="-20" y1="35" x2="110" y2="35" stroke={corTorre} strokeWidth="2.5"/>
        <line x1="-12" y1="12" x2="102" y2="12" stroke={corTorre} strokeWidth="2.5"/>
        <line x1="45" y1="0" x2="45" y2="12" stroke={corTorre} strokeWidth="2.5"/>
        <circle cx="-20" cy="35" r="3.5" fill={corTorre}/>
        <circle cx="110" cy="35" r="3.5" fill={corTorre}/>
        <circle cx="-12" cy="12" r="3.5" fill={corTorre}/>
        <circle cx="102" cy="12" r="3.5" fill={corTorre}/>
        <line x1="5" y1="348" x2="22" y2="360" stroke={corTorre} strokeWidth="2"/>
        <line x1="85" y1="348" x2="68" y2="360" stroke={corTorre} strokeWidth="2"/>
      </g>

      {/* TORRE 3 — grande central */}
      <g transform="translate(560,0)">
        <line x1="55" y1="0" x2="25" y2="400" stroke="url(#torreGrad)" strokeWidth="3.5"/>
        <line x1="55" y1="0" x2="85" y2="400" stroke="url(#torreGrad)" strokeWidth="3.5"/>
        <line x1="25" y1="400" x2="85" y2="400" stroke={corTorre} strokeWidth="3.5"/>
        <line x1="25" y1="400" x2="55" y2="280" stroke={corTorre} strokeWidth="1.6"/>
        <line x1="85" y1="400" x2="55" y2="280" stroke={corTorre} strokeWidth="1.6"/>
        <line x1="25" y1="280" x2="55" y2="180" stroke={corTorre} strokeWidth="1.6"/>
        <line x1="85" y1="280" x2="55" y2="180" stroke={corTorre} strokeWidth="1.6"/>
        <line x1="28" y1="180" x2="55" y2="100" stroke={corTorre} strokeWidth="1.6"/>
        <line x1="82" y1="180" x2="55" y2="100" stroke={corTorre} strokeWidth="1.6"/>
        <line x1="32" y1="100" x2="55" y2="45" stroke={corTorre} strokeWidth="1.6"/>
        <line x1="78" y1="100" x2="55" y2="45" stroke={corTorre} strokeWidth="1.6"/>
        <line x1="5" y1="280" x2="105" y2="280" stroke={corTorre} strokeWidth="2.2"/>
        <line x1="8" y1="180" x2="102" y2="180" stroke={corTorre} strokeWidth="2.2"/>
        <line x1="12" y1="100" x2="98" y2="100" stroke={corTorre} strokeWidth="2.2"/>
        <line x1="16" y1="45" x2="94" y2="45" stroke={corTorre} strokeWidth="2.5"/>
        <line x1="-30" y1="18" x2="140" y2="18" stroke={corTorre} strokeWidth="3"/>
        <line x1="-18" y1="5" x2="128" y2="5" stroke={corTorre} strokeWidth="3"/>
        <line x1="55" y1="0" x2="55" y2="5" stroke={corTorre} strokeWidth="3"/>
        <circle cx="-30" cy="18" r="4.5" fill={corTorre}/>
        <circle cx="140" cy="18" r="4.5" fill={corTorre}/>
        <circle cx="-18" cy="5" r="4" fill={corTorre}/>
        <circle cx="128" cy="5" r="4" fill={corTorre}/>
        <line x1="5" y1="385" x2="25" y2="400" stroke={corTorre} strokeWidth="2.5"/>
        <line x1="105" y1="385" x2="85" y2="400" stroke={corTorre} strokeWidth="2.5"/>
      </g>

      {/* TORRE 4 — média direita */}
      <g transform="translate(870,50)">
        <line x1="45" y1="0" x2="22" y2="350" stroke="url(#torreGrad)" strokeWidth="3"/>
        <line x1="45" y1="0" x2="68" y2="350" stroke="url(#torreGrad)" strokeWidth="3"/>
        <line x1="22" y1="350" x2="68" y2="350" stroke={corTorre} strokeWidth="3"/>
        <line x1="22" y1="350" x2="45" y2="230" stroke={corTorre} strokeWidth="1.4"/>
        <line x1="68" y1="350" x2="45" y2="230" stroke={corTorre} strokeWidth="1.4"/>
        <line x1="22" y1="230" x2="45" y2="140" stroke={corTorre} strokeWidth="1.4"/>
        <line x1="68" y1="230" x2="45" y2="140" stroke={corTorre} strokeWidth="1.4"/>
        <line x1="26" y1="140" x2="45" y2="70" stroke={corTorre} strokeWidth="1.4"/>
        <line x1="64" y1="140" x2="45" y2="70" stroke={corTorre} strokeWidth="1.4"/>
        <line x1="8" y1="230" x2="82" y2="230" stroke={corTorre} strokeWidth="2"/>
        <line x1="10" y1="140" x2="80" y2="140" stroke={corTorre} strokeWidth="2"/>
        <line x1="13" y1="70" x2="77" y2="70" stroke={corTorre} strokeWidth="2.2"/>
        <line x1="-20" y1="32" x2="110" y2="32" stroke={corTorre} strokeWidth="2.5"/>
        <line x1="-12" y1="10" x2="102" y2="10" stroke={corTorre} strokeWidth="2.5"/>
        <line x1="45" y1="0" x2="45" y2="10" stroke={corTorre} strokeWidth="2.5"/>
        <circle cx="-20" cy="32" r="3.5" fill={corTorre}/>
        <circle cx="110" cy="32" r="3.5" fill={corTorre}/>
        <circle cx="-12" cy="10" r="3.5" fill={corTorre}/>
        <circle cx="102" cy="10" r="3.5" fill={corTorre}/>
        <line x1="5" y1="338" x2="22" y2="350" stroke={corTorre} strokeWidth="2"/>
        <line x1="85" y1="338" x2="68" y2="350" stroke={corTorre} strokeWidth="2"/>
      </g>

      {/* TORRE 5 — pequena direita */}
      <g transform="translate(1100,100)">
        <line x1="35" y1="0" x2="18" y2="300" stroke="url(#torreGrad)" strokeWidth="2.5"/>
        <line x1="35" y1="0" x2="52" y2="300" stroke="url(#torreGrad)" strokeWidth="2.5"/>
        <line x1="18" y1="300" x2="52" y2="300" stroke={corTorre} strokeWidth="2.5"/>
        <line x1="18" y1="300" x2="35" y2="190" stroke={corTorre} strokeWidth="1.2"/>
        <line x1="52" y1="300" x2="35" y2="190" stroke={corTorre} strokeWidth="1.2"/>
        <line x1="18" y1="190" x2="35" y2="110" stroke={corTorre} strokeWidth="1.2"/>
        <line x1="52" y1="190" x2="35" y2="110" stroke={corTorre} strokeWidth="1.2"/>
        <line x1="10" y1="190" x2="60" y2="190" stroke={corTorre} strokeWidth="1.8"/>
        <line x1="12" y1="110" x2="58" y2="110" stroke={corTorre} strokeWidth="1.8"/>
        <line x1="-15" y1="26" x2="85" y2="26" stroke={corTorre} strokeWidth="2.2"/>
        <line x1="-8" y1="8" x2="78" y2="8" stroke={corTorre} strokeWidth="2.2"/>
        <line x1="35" y1="0" x2="35" y2="8" stroke={corTorre} strokeWidth="2"/>
        <circle cx="-15" cy="26" r="3" fill={corTorre}/>
        <circle cx="85" cy="26" r="3" fill={corTorre}/>
      </g>

      {/* CABOS DE TRANSMISSÃO entre torres */}
      <path d="M 115 10 Q 200 55 260 12" stroke={corTorre} strokeWidth="1.2" fill="none" opacity="0.8"/>
      <path d="M 115 28 Q 200 75 260 35" stroke={corTorre} strokeWidth="1.2" fill="none" opacity="0.8"/>
      <path d="M 390 12 Q 475 45 530 18" stroke={corTorre} strokeWidth="1.3" fill="none" opacity="0.8"/>
      <path d="M 390 35 Q 475 68 530 45" stroke={corTorre} strokeWidth="1.3" fill="none" opacity="0.8"/>
      <path d="M 700 18 Q 785 50 850 32" stroke={corTorre} strokeWidth="1.3" fill="none" opacity="0.8"/>
      <path d="M 700 45 Q 785 78 850 62" stroke={corTorre} strokeWidth="1.3" fill="none" opacity="0.8"/>
      <path d="M 980 32 Q 1040 55 1085 26" stroke={corTorre} strokeWidth="1.2" fill="none" opacity="0.8"/>
      <path d="M 980 62 Q 1040 82 1085 55" stroke={corTorre} strokeWidth="1.2" fill="none" opacity="0.8"/>

      {/* linha do horizonte / chão */}
      <line x1="0" y1="398" x2="1200" y2="398" stroke={corTorre} strokeWidth="1" opacity="0.3"/>
    </svg>
  );
}

// ============================================================
// TELA: RELATÓRIOS
// ============================================================
function Relatorios({ notas }) {
  const [periodo, setPeriodo] = useState("mensal");
  const [empresa, setEmpresa] = useState("");
  const [status, setStatus] = useState("");
  const [dias, setDias] = useState("");
  const [valorFiltro, setValorFiltro] = useState("");
  const [tipoData, setTipoData] = useState("emissao");
  const [dtDe, setDtDe] = useState("");
  const [dtAte, setDtAte] = useState("");
  const [finalidade, setFinalidade] = useState("");
  const [busca, setBusca] = useState("");

  function parseDateBR2(str) {
    if (!str || str.length !== 10) return null;
    const [d, m, y] = str.split("/");
    if (!d || !m || !y) return null;
    return new Date(`${y}-${m}-${d}`);
  }

  const filtradas = notas.filter(n => {
    if (empresa && n.empresa !== empresa) return false;
    if (status && n.status !== status) return false;
    if (finalidade && n.finalidade !== finalidade) return false;
    if (dias === "critico" && n.qtdeDias < 60) return false;
    if (dias === "atencao" && (n.qtdeDias < 25 || n.qtdeDias >= 60)) return false;
    if (dias === "ok" && n.qtdeDias >= 25) return false;
    if (valorFiltro === "acima25k" && n.valor <= 25000) return false;
    if (valorFiltro === "ate25k" && n.valor > 25000) return false;
    if (busca) {
      const b = busca.toLowerCase();
      if (!n.razaoSocial.toLowerCase().includes(b) && !n.numNota.includes(b) && !n.fornecedor.includes(b)) return false;
    }
    if (dtDe || dtAte) {
      const campoData = tipoData === "emissao" ? n.dtEmissao : n.dtImportacao;
      const dataRef = parseDateBR2(campoData);
      const de = parseDateBR2(dtDe);
      const ate = parseDateBR2(dtAte);
      if (!dataRef) return false;
      if (de && dataRef < de) return false;
      if (ate && dataRef > ate) return false;
    }
    return true;
  });

  const totalTaxas = filtradas.reduce((s, n) => s + (n.taxaReanalise || 0) + (n.taxaDesembaraco || 0), 0);
  const totalICMS = filtradas.reduce((s, n) => s + (n.icmsAntecipado || 0), 0);
  const totalMultas = filtradas.reduce((s, n) => s + (n.multa || 0), 0);
  const totalJuros = filtradas.reduce((s, n) => s + (n.juros || 0), 0);
  const totalGeral = totalTaxas + totalICMS + totalMultas + totalJuros;

  const porStatus = STATUS_LIST.map(s => ({ status: s, count: filtradas.filter(n => n.status === s).length })).filter(x => x.count > 0);
  const porEmpresa = EMPRESAS.map(e => ({ ...e, count: filtradas.filter(n => n.empresa === e.id).length, custos: filtradas.filter(n => n.empresa === e.id).reduce((s, n) => s + (n.taxaReanalise || 0) + (n.taxaDesembaraco || 0) + (n.icmsAntecipado || 0) + (n.multa10pct || 0) + (n.multa || 0) + (n.juros || 0), 0) }));
  const porFinalidade = FINALIDADES.map(f => ({ f, count: filtradas.filter(n => n.finalidade === f).length })).filter(x => x.count > 0);

  const temFiltro = empresa || status || dias || dtDe || dtAte || finalidade || busca || valorFiltro;
  const limpar = () => { setEmpresa(""); setStatus(""); setDias(""); setDtDe(""); setDtAte(""); setFinalidade(""); setBusca(""); setTipoData("emissao"); setValorFiltro(""); };

  function fmtDateInput(setter) {
    return (e) => {
      let v = e.target.value.replace(/\D/g, "");
      if (v.length >= 3) v = v.slice(0,2) + "/" + v.slice(2);
      if (v.length >= 6) v = v.slice(0,5) + "/" + v.slice(5);
      setter(v.slice(0,10));
    };
  }

  return (
    <div className="space-y-5">
      {/* Filtros */}
      <div className="rounded-2xl border p-4 space-y-3" style={{ borderColor: "#f0f0f0", background: "#fafafa" }}>
        <div className="flex items-center justify-between">
          <p className="text-xs font-bold text-gray-500 uppercase tracking-wide">Filtros do Relatório</p>
          <div className="flex items-center gap-3">
            {temFiltro && <button onClick={limpar} className="text-xs font-semibold" style={{ color: "#E8450A" }}>✕ Limpar</button>}
            <div className="flex gap-2">
              <button className="px-4 py-2 rounded-lg text-sm font-semibold border" style={{ borderColor: "#E8450A", color: "#E8450A" }}
                onClick={() => exportarExcel(filtradas, EMPRESAS, periodo)}>📊 Excel</button>
              <button className="px-4 py-2 rounded-lg text-sm font-semibold text-white" style={{ background: "#E8450A" }}
                onClick={() => exportarPDF(filtradas, EMPRESAS, periodo)}>📄 PDF</button>
            </div>
          </div>
        </div>

        {/* Linha 1 */}
        <div className="flex flex-wrap gap-2">
          <select value={periodo} onChange={e => setPeriodo(e.target.value)} className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: "#e5e7eb" }}>
            <option value="semanal">Semanal</option>
            <option value="mensal">Mensal</option>
            <option value="trimestral">Trimestral</option>
          </select>
          <input placeholder="🔍 Buscar fornecedor ou NF..." value={busca} onChange={e => setBusca(e.target.value)}
            className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: "#e5e7eb", minWidth: 200 }} />
          <select value={empresa} onChange={e => setEmpresa(e.target.value)} className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: "#e5e7eb" }}>
            <option value="">Todas as empresas</option>
            {EMPRESAS.map(e => <option key={e.id} value={e.id}>{e.nome}</option>)}
          </select>
          <select value={status} onChange={e => setStatus(e.target.value)} className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: "#e5e7eb" }}>
            <option value="">Todos os status</option>
            {STATUS_LIST.map(s => <option key={s}>{s}</option>)}
          </select>
          <select value={finalidade} onChange={e => setFinalidade(e.target.value)} className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: "#e5e7eb" }}>
            <option value="">Todas as finalidades</option>
            {FINALIDADES.map(f => <option key={f}>{f}</option>)}
          </select>
          <select value={dias} onChange={e => setDias(e.target.value)} className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: "#e5e7eb" }}>
            <option value="">Todos os prazos</option>
            <option value="critico">🔴 Crítico (60+)</option>
            <option value="atencao">🟡 Atenção (25-59)</option>
            <option value="ok">🟢 OK (0-24)</option>
          </select>
          <select value={valorFiltro} onChange={e => setValorFiltro(e.target.value)} className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: valorFiltro ? "#7e3af2" : "#e5e7eb", color: valorFiltro === "acima25k" ? "#7e3af2" : undefined }}>
            <option value="">Todos os valores</option>
            <option value="acima25k">⚠️ Acima de R$ 25k</option>
            <option value="ate25k">Até R$ 25k</option>
          </select>
        </div>

        {/* Linha 2: datas */}
        <div className="flex flex-wrap items-center gap-2">
          <p className="text-xs font-semibold text-gray-500">Período:</p>
          <select value={tipoData} onChange={e => setTipoData(e.target.value)} className="border rounded-lg px-3 py-2 text-sm bg-white" style={{ borderColor: "#e5e7eb" }}>
            <option value="emissao">Data de Emissão</option>
            <option value="importacao">Data de Importação</option>
          </select>
          <span className="text-xs text-gray-400">De</span>
          <input type="text" placeholder="DD/MM/AAAA" value={dtDe} maxLength={10} onChange={fmtDateInput(setDtDe)}
            className="border rounded-lg px-3 py-2 text-sm bg-white w-32" style={{ borderColor: dtDe.length===10 ? "#E8450A" : "#e5e7eb" }} />
          <span className="text-xs text-gray-400">Até</span>
          <input type="text" placeholder="DD/MM/AAAA" value={dtAte} maxLength={10} onChange={fmtDateInput(setDtAte)}
            className="border rounded-lg px-3 py-2 text-sm bg-white w-32" style={{ borderColor: dtAte.length===10 ? "#E8450A" : "#e5e7eb" }} />
          <span className="text-xs text-gray-400 ml-2">{filtradas.length} nota(s) no relatório{temFiltro && <span className="font-semibold" style={{ color: "#E8450A" }}> • filtros ativos</span>}</span>
        </div>
      </div>

      {/* Relatório */}
      <div className="rounded-2xl border p-6" style={{ borderColor: "#f0f0f0" }}>
        <div className="flex items-center justify-between mb-1">
          <h3 className="font-bold text-gray-700 uppercase tracking-wide text-sm">Relatório Gerencial — {periodo.charAt(0).toUpperCase() + periodo.slice(1)}</h3>
        </div>
        <p className="text-xs text-gray-400 mb-5">{empresa ? EMPRESAS.find(e => e.id === empresa)?.nome : "Todas as empresas"} {dtDe && `• De ${dtDe}`} {dtAte && `Até ${dtAte}`}</p>

        <div className="grid grid-cols-2 gap-4 mb-6" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))" }}>
          {[
            { label: "Total de Notas", valor: filtradas.length, cor: "#374151" },
            { label: "Notas Ativas", valor: filtradas.filter(n => !["Desembaraçada", "Recusada"].includes(n.status)).length, cor: "#1a56db" },
            { label: "Desembaraçadas", valor: filtradas.filter(n => n.status === "Desembaraçada").length, cor: "#2d6a4f" },
            { label: "Recusadas", valor: filtradas.filter(n => n.status === "Recusada").length, cor: "#c0392b" },
            { label: "Críticas (60+d)", valor: filtradas.filter(n => n.qtdeDias >= 60 && !["Desembaraçada","Recusada"].includes(n.status)).length, cor: "#c0392b" },
            { label: "Acima R$25k", valor: filtradas.filter(n => n.valor > 25000).length, cor: "#7e3af2" },
          ].map((c, i) => (
            <div key={i} className="p-4 rounded-xl" style={{ background: "#f8f9fa" }}>
              <p className="text-2xl font-black" style={{ color: c.cor }}>{c.valor}</p>
              <p className="text-xs text-gray-500 mt-1">{c.label}</p>
            </div>
          ))}
        </div>

        <h4 className="font-semibold text-gray-600 mb-3 text-sm">Breakdown de Custos</h4>
        <div className="space-y-2 mb-6">
          {[
            { label: "Taxas (Reanálise + Desembaraço)", valor: totalTaxas },
            { label: "ICMS Antecipado", valor: totalICMS },
            { label: "Multas", valor: totalMultas },
            { label: "Juros", valor: totalJuros },
          ].map((c, i) => (
            <div key={i} className="flex justify-between items-center p-3 rounded-lg" style={{ background: "#f8f9fa" }}>
              <span className="text-sm text-gray-600">{c.label}</span>
              <span className="font-bold text-sm" style={{ color: c.valor > 0 ? "#c0392b" : "#9ca3af" }}>{fmtMoeda(c.valor)}</span>
            </div>
          ))}
          <div className="flex justify-between items-center p-3 rounded-xl" style={{ background: "#f0f8f8", border: "1px solid #ffd6b8" }}>
            <span className="font-bold text-gray-700">TOTAL GERAL</span>
            <span className="font-black text-xl" style={{ color: "#E8450A" }}>{fmtMoeda(totalGeral)}</span>
          </div>
        </div>

        <div className="grid grid-cols-1 gap-6" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(280px,1fr))" }}>
          <div>
            <h4 className="font-semibold text-gray-600 mb-3 text-sm">Por Empresa</h4>
            <div className="space-y-2">
              {porEmpresa.map(e => (
                <div key={e.id} className="flex justify-between items-center p-3 rounded-lg" style={{ background: "#f8f9fa" }}>
                  <span className="text-sm text-gray-700">{e.nome}</span>
                  <div className="text-right">
                    <span className="font-semibold text-sm text-gray-700">{e.count} notas</span>
                    <p className="text-xs" style={{ color: e.custos > 0 ? "#c0392b" : "#9ca3af" }}>{fmtMoeda(e.custos)}</p>
                  </div>
                </div>
              ))}
            </div>
          </div>
          <div>
            <h4 className="font-semibold text-gray-600 mb-3 text-sm">Por Status</h4>
            <div className="space-y-2">
              {porStatus.map(s => (
                <div key={s.status} className="flex items-center justify-between px-3 py-2 rounded-lg" style={{ background: "#f8f9fa" }}>
                  <Badge status={s.status} />
                  <span className="font-bold text-sm text-gray-700">{s.count}</span>
                </div>
              ))}
            </div>
          </div>
          {porFinalidade.length > 0 && (
            <div>
              <h4 className="font-semibold text-gray-600 mb-3 text-sm">Por Finalidade</h4>
              <div className="space-y-2">
                {porFinalidade.map(({ f, count }) => (
                  <div key={f} className="flex items-center justify-between px-3 py-2 rounded-lg" style={{ background: "#f8f9fa" }}>
                    <span className="text-sm text-gray-600">{f}</span>
                    <span className="font-bold text-sm text-gray-700">{count}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ============================================================
// MODAL EMPRESA
// ============================================================
function ModalEmpresa({ empresa, onClose, onSalvar }) {
  const [form, setForm] = useState(empresa || { id: "", nome: "", inscricao: "", cnpj: "", ativa: true });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const isNova = !empresa;

  const salvar = () => {
    if (!form.nome || !form.inscricao || !form.cnpj) { alert("Preencha Nome, IE e CNPJ."); return; }
    const id = isNova ? "emp_" + Date.now() : form.id;
    onSalvar({ ...form, id });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background: "rgba(0,0,0,0.5)" }}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md m-4">
        <div className="flex items-center justify-between p-5 border-b" style={{ borderColor: "#f0f0f0" }}>
          <h2 className="font-bold text-gray-800">{isNova ? "Nova Empresa" : "Editar Empresa"}</h2>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 text-2xl font-light">×</button>
        </div>
        <div className="p-5 space-y-4">
          <div>
            <label className="text-xs font-semibold text-gray-500 uppercase">Nome / Apelido da Filial</label>
            <input value={form.nome} onChange={e => set("nome", e.target.value)} placeholder="Ex: FILIAL 02 - MANAUS"
              className="mt-1 w-full border rounded-lg px-3 py-2 text-sm" style={{ borderColor: "#e5e7eb" }} />
          </div>
          <div>
            <label className="text-xs font-semibold text-gray-500 uppercase">Inscrição Estadual (IE)</label>
            <input value={form.inscricao} onChange={e => set("inscricao", e.target.value)} placeholder="Ex: 04.235.429-3"
              className="mt-1 w-full border rounded-lg px-3 py-2 text-sm" style={{ borderColor: "#e5e7eb" }} />
            <p className="text-xs text-gray-400 mt-1">Usada para identificação automática no upload do CSV.</p>
          </div>
          <div>
            <label className="text-xs font-semibold text-gray-500 uppercase">CNPJ</label>
            <input value={form.cnpj} onChange={e => set("cnpj", e.target.value)} placeholder="Ex: 07.791.042/0002-18"
              className="mt-1 w-full border rounded-lg px-3 py-2 text-sm" style={{ borderColor: "#e5e7eb" }} />
            <p className="text-xs text-gray-400 mt-1">Também usado para identificação automática no upload.</p>
          </div>
          <div className="flex items-center gap-3 p-3 rounded-xl" style={{ background: "#f8f9fa" }}>
            <span className="text-sm text-gray-700 flex-1">Empresa Ativa</span>
            <button onClick={() => set("ativa", !form.ativa)}
              className="w-12 h-6 rounded-full transition-all relative"
              style={{ background: form.ativa ? "#E8450A" : "#d1d5db" }}>
              <span className="absolute top-1 w-4 h-4 bg-white rounded-full shadow transition-all"
                style={{ left: form.ativa ? "26px" : "2px" }} />
            </button>
            <span className="text-xs font-semibold" style={{ color: form.ativa ? "#E8450A" : "#9ca3af" }}>
              {form.ativa ? "Ativa" : "Inativa"}
            </span>
          </div>
        </div>
        <div className="flex justify-end gap-3 p-5 border-t" style={{ borderColor: "#f0f0f0" }}>
          <button onClick={onClose} className="px-4 py-2 rounded-lg border text-sm text-gray-600 hover:bg-gray-50" style={{ borderColor: "#e5e7eb" }}>Cancelar</button>
          <button onClick={salvar} className="px-6 py-2 rounded-lg text-sm font-semibold text-white" style={{ background: "#E8450A" }}>
            {isNova ? "Adicionar Empresa" : "Salvar Alterações"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// MODAL CONFIRMAÇÃO EMPRESA NO IMPORT
// ============================================================
function ModalConfirmImport({ notasParaImportar, empresas, onConfirmar, onCancelar, usuarioNome, fileName }) {
  // agrupa por CNPJ detectado do arquivo
  const grupos = [];
  const cnpjsVistos = {};
  notasParaImportar.forEach(n => {
    const cnpj = n._cnpjDetectado || "desconhecido";
    if (!cnpjsVistos[cnpj]) {
      cnpjsVistos[cnpj] = { cnpj, notas: [], empresaId: n._empresaId || "" };
      grupos.push(cnpjsVistos[cnpj]);
    }
    cnpjsVistos[cnpj].notas.push(n);
  });

  const [selecoes, setSelecoes] = useState(() => {
    const s = {};
    grupos.forEach(g => { s[g.cnpj] = g.empresaId; });
    return s;
  });
  const [empresaGlobal, setEmpresaGlobal] = useState("");

  // Quando usuário escolhe empresa global, aplica a todos os grupos
  const aplicarGlobal = (empId) => {
    setEmpresaGlobal(empId);
    if (empId) {
      const novas = {};
      grupos.forEach(g => { novas[g.cnpj] = empId; });
      setSelecoes(novas);
    }
  };

  const confirmar = () => {
    const notasFinais = notasParaImportar.map(n => ({
      ...n,
      empresa: selecoes[n._cnpjDetectado || "desconhecido"] || "filial02",
      _cnpjDetectado: undefined,
      _empresaId: undefined,
      historico: [{ acao: `Importada via arquivo: ${fileName}`, usuario: usuarioNome, data: new Date().toLocaleString("pt-BR") }]
    }));
    onConfirmar(notasFinais);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background: "rgba(0,0,0,0.6)" }}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg m-4 max-h-screen overflow-y-auto">
        <div className="p-5 border-b" style={{ borderColor: "#f0f0f0" }}>
          <h2 className="font-bold text-gray-800">Confirmar Importação</h2>
          <p className="text-xs text-gray-400 mt-1">{notasParaImportar.length} nota(s) novas encontradas em <span className="font-semibold">{fileName}</span></p>
        </div>

        <div className="p-5 space-y-4">

          {/* TAG: Aplicar mesma empresa a todas */}
          <div className="p-4 rounded-xl" style={{ background: "#f0f8f8", border: "2px solid #E8450A" }}>
            <div className="flex items-center gap-2 mb-2">
              <span className="text-xs font-black px-2 py-1 rounded-full text-white" style={{ background: "#E8450A" }}>⚡ ATALHO</span>
              <p className="text-sm font-bold" style={{ color: "#1a4a4a" }}>Aplicar a mesma empresa em todas as notas</p>
            </div>
            <select
              value={empresaGlobal}
              onChange={e => aplicarGlobal(e.target.value)}
              className="w-full border-2 rounded-lg px-3 py-2 text-sm font-semibold"
              style={{ borderColor: empresaGlobal ? "#E8450A" : "#d0e8e8", color: empresaGlobal ? "#1a4a4a" : "#888" }}>
              <option value="">Selecione para aplicar a todas...</option>
              {empresas.filter(e => e.ativa).map(e => (
                <option key={e.id} value={e.id}>{e.nome} — IE: {e.inscricao}</option>
              ))}
            </select>
            {empresaGlobal && (
              <p className="text-xs mt-2 font-semibold" style={{ color: "#E8450A" }}>
                ✅ {empresas.find(e => e.id === empresaGlobal)?.nome} aplicada a todas as {notasParaImportar.length} notas
              </p>
            )}
          </div>

          {/* Divisor */}
          <div className="flex items-center gap-3">
            <div className="flex-1 h-px" style={{ background: "#e5e7eb" }} />
            <span className="text-xs text-gray-400 font-semibold">ou ajuste individualmente por CNPJ</span>
            <div className="flex-1 h-px" style={{ background: "#e5e7eb" }} />
          </div>

          {/* Grupos por CNPJ */}
          {grupos.map(g => (
            <div key={g.cnpj} className="p-4 rounded-xl border" style={{ borderColor: selecoes[g.cnpj] ? "#d1fae5" : "#fde68a", background: selecoes[g.cnpj] ? "#f0fdf4" : "#fffbeb" }}>
              <div className="flex items-center gap-2 mb-2">
                <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: selecoes[g.cnpj] ? "#d1fae5" : "#fde68a", color: selecoes[g.cnpj] ? "#065f46" : "#92400e" }}>
                  {selecoes[g.cnpj] ? "✅ Definida" : "⚠️ Pendente"}
                </span>
                <span className="text-xs text-gray-500 font-mono">{g.cnpj}</span>
                <span className="text-xs text-gray-400">({g.notas.length} nota{g.notas.length > 1 ? "s" : ""})</span>
              </div>
              <label className="text-xs font-semibold text-gray-500 uppercase">Empresa</label>
              <select value={selecoes[g.cnpj] || ""} onChange={e => { setEmpresaGlobal(""); setSelecoes(s => ({ ...s, [g.cnpj]: e.target.value })); }}
                className="mt-1 w-full border rounded-lg px-3 py-2 text-sm" style={{ borderColor: "#e5e7eb" }}>
                <option value="">Selecione a empresa...</option>
                {empresas.filter(e => e.ativa).map(e => (
                  <option key={e.id} value={e.id}>{e.nome} — IE: {e.inscricao}</option>
                ))}
              </select>
            </div>
          ))}
        </div>

        <div className="flex justify-end gap-3 p-5 border-t" style={{ borderColor: "#f0f0f0" }}>
          <button onClick={onCancelar} className="px-4 py-2 rounded-lg border text-sm text-gray-600 hover:bg-gray-50" style={{ borderColor: "#e5e7eb" }}>Cancelar</button>
          <button onClick={confirmar}
            disabled={grupos.some(g => !selecoes[g.cnpj])}
            className="px-6 py-2 rounded-lg text-sm font-semibold text-white disabled:opacity-50"
            style={{ background: "#E8450A" }}>
            Importar {notasParaImportar.length} Nota(s)
          </button>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// TELA: CONFIGURAÇÕES
// ============================================================
// ============================================================
// ABA FINANCEIRO — LANÇAMENTOS COM MODAL
// ============================================================
const TIPOS_LANCAMENTO = ["Taxa de Reanálise","Taxa de Desembaraço","ICMS Antecipado","Multa 10%","Multa Adicional","Juros","Outro"];

function AbaFinanceiro({ form, set, fmtMoeda }) {
  const lancamentos = form.lancamentos || [];
  const [modalAberto, setModalAberto] = useState(false);
  const [novoLanc, setNovoLanc] = useState({ tipo:"Taxa de Desembaraço", codigo:"", valor:"", venc:"" });
  const [confirmExcluir, setConfirmExcluir] = useState(null);

  function fmtDate(val) {
    let v = val.replace(/\D/g,"");
    if(v.length>=3) v=v.slice(0,2)+"/"+v.slice(2);
    if(v.length>=6) v=v.slice(0,5)+"/"+v.slice(5);
    return v.slice(0,10);
  }
  function salvarLanc() {
    if(!novoLanc.valor) return;
    set("lancamentos", [...lancamentos, { id: Date.now(), ...novoLanc, pago: false }]);
    setNovoLanc({ tipo:"Taxa de Desembaraço", codigo:"", valor:"", venc:"" });
    setModalAberto(false);
  }
  function togglePago(idx) {
    set("lancamentos", lancamentos.map((l,i) => i===idx ? {...l, pago:!l.pago} : l));
  }
  function excluir(idx) {
    set("lancamentos", lancamentos.filter((_,i) => i!==idx));
    setConfirmExcluir(null);
  }

  const totalGeral = lancamentos.reduce((s,l)=>s+(parseFloat(l.valor)||0),0);
  const pago = lancamentos.filter(l=>l.pago).reduce((s,l)=>s+(parseFloat(l.valor)||0),0);
  const pendente = totalGeral - pago;

  return (
    <div className="space-y-3">
      {/* Header */}
      <div className="flex items-center justify-between">
        <p className="text-xs text-gray-400">Taxas, ICMS, multas e juros desta nota.</p>
        <button onClick={() => setModalAberto(true)}
          className="px-4 py-1.5 rounded-lg text-xs font-bold text-white"
          style={{ background:"#1a4a4a" }}>
          + Novo lançamento
        </button>
      </div>

      {/* Lista */}
      {lancamentos.length === 0 && (
        <div className="text-center py-8 text-gray-300 rounded-xl border border-dashed" style={{ borderColor:"#e5e7eb" }}>
          <p className="text-sm">Nenhum lançamento registrado</p>
          <p className="text-xs mt-1">Clique em "+ Novo lançamento" para começar</p>
        </div>
      )}

      {lancamentos.map((l, idx) => (
        <div key={l.id||idx} className="flex items-center gap-3 px-4 py-3 rounded-xl"
          style={{ background: l.pago ? "#f0fdf4" : "#f8f9fa", border: l.pago ? "1px solid #86efac" : "1px solid #e5e7eb" }}>
          <div className="flex-1 min-w-0">
            <p className="text-sm font-semibold text-gray-800">{l.tipo}</p>
            <p className="text-xs text-gray-400 mt-0.5">
              {l.codigo && <span className="mr-2">Cód: {l.codigo}</span>}
              {l.venc && <span>Venc: {l.venc}</span>}
            </p>
          </div>
          <span className="text-sm font-bold" style={{ color: l.pago ? "#2d6a4f" : "#374151" }}>{fmtMoeda(parseFloat(l.valor)||0)}</span>
          <span className="text-xs font-semibold px-2 py-1 rounded-full flex-shrink-0"
            style={{ background: l.pago ? "#dcfce7" : "#fef9c3", color: l.pago ? "#166534" : "#854d0e" }}>
            {l.pago ? "Pago" : "Pendente"}
          </span>
          <label className="flex items-center gap-1 cursor-pointer flex-shrink-0">
            <input type="checkbox" checked={l.pago||false} onChange={()=>togglePago(idx)} className="w-4 h-4 accent-green-600" />
            <span className="text-xs text-gray-400">Pago</span>
          </label>
          <button onClick={() => setConfirmExcluir(idx)}
            className="text-gray-300 hover:text-red-400 transition-colors text-lg leading-none flex-shrink-0">🗑</button>
        </div>
      ))}

      {/* Total */}
      {lancamentos.length > 0 && (
        <div className="p-4 rounded-xl" style={{ background:"#1a4a4a" }}>
          <div className="flex items-center justify-between">
            <div>
              <p className="text-xs font-semibold uppercase tracking-wide" style={{ color:"#7ecece" }}>Total Geral</p>
              <p className="text-2xl font-black text-white mt-1">{fmtMoeda(totalGeral)}</p>
            </div>
            <div className="text-right space-y-1">
              <p className="text-xs" style={{ color:"#7ecece" }}>Pago: <span className="font-semibold text-white">{fmtMoeda(pago)}</span></p>
              <p className="text-xs" style={{ color:"#f0a500" }}>Pendente: <span className="font-semibold">{fmtMoeda(pendente)}</span></p>
            </div>
          </div>
        </div>
      )}

      {/* Modal novo lançamento */}
      {modalAberto && (
        <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background:"rgba(0,0,0,0.5)" }}>
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-sm m-4">
            <div className="flex items-center justify-between p-5 border-b" style={{ borderColor:"#f0f0f0" }}>
              <h2 className="font-bold text-gray-800">Novo Lançamento</h2>
              <button onClick={()=>setModalAberto(false)} className="text-gray-400 hover:text-gray-600 text-2xl font-light">×</button>
            </div>
            <div className="p-5 space-y-4">
              <div>
                <label className="text-xs font-semibold text-gray-500 uppercase">Tipo</label>
                <select value={novoLanc.tipo} onChange={e=>setNovoLanc(l=>({...l,tipo:e.target.value}))}
                  className="mt-1 w-full border rounded-lg p-2.5 text-sm" style={{ borderColor:"#e5e7eb" }}>
                  {TIPOS_LANCAMENTO.map(t=><option key={t}>{t}</option>)}
                </select>
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs font-semibold text-gray-500 uppercase">Código</label>
                  <input value={novoLanc.codigo} onChange={e=>setNovoLanc(l=>({...l,codigo:e.target.value}))}
                    placeholder="Ex: 1.1-IM"
                    className="mt-1 w-full border rounded-lg p-2.5 text-sm" style={{ borderColor:"#e5e7eb" }} />
                </div>
                <div>
                  <label className="text-xs font-semibold text-gray-500 uppercase">Valor (R$)</label>
                  <input type="number" step="0.01" value={novoLanc.valor} onChange={e=>setNovoLanc(l=>({...l,valor:e.target.value}))}
                    placeholder="0,00"
                    className="mt-1 w-full border rounded-lg p-2.5 text-sm" style={{ borderColor:"#e5e7eb" }} />
                </div>
              </div>
              <div>
                <label className="text-xs font-semibold text-gray-500 uppercase">Vencimento</label>
                <input value={novoLanc.venc} onChange={e=>setNovoLanc(l=>({...l,venc:fmtDate(e.target.value)}))}
                  placeholder="DD/MM/AAAA" maxLength={10}
                  className="mt-1 w-full border rounded-lg p-2.5 text-sm" style={{ borderColor: novoLanc.venc?.length===10?"#E8450A":"#e5e7eb" }} />
              </div>
            </div>
            <div className="flex justify-end gap-3 p-5 border-t" style={{ borderColor:"#f0f0f0" }}>
              <button onClick={()=>setModalAberto(false)} className="px-4 py-2 rounded-lg border text-sm text-gray-600" style={{ borderColor:"#e5e7eb" }}>Cancelar</button>
              <button onClick={salvarLanc} disabled={!novoLanc.valor}
                className="px-6 py-2 rounded-lg text-sm font-semibold text-white disabled:opacity-50"
                style={{ background:"#1a4a4a" }}>
                Salvar lançamento
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal confirmar exclusão */}
      {confirmExcluir !== null && (
        <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background:"rgba(0,0,0,0.5)" }}>
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-xs m-4 p-6 text-center">
            <p className="text-2xl mb-3">🗑</p>
            <p className="font-bold text-gray-800 mb-1">Excluir lançamento?</p>
            <p className="text-sm text-gray-400 mb-5">Esta ação não pode ser desfeita.</p>
            <div className="flex gap-3 justify-center">
              <button onClick={()=>setConfirmExcluir(null)} className="px-4 py-2 rounded-lg border text-sm text-gray-600" style={{ borderColor:"#e5e7eb" }}>Cancelar</button>
              <button onClick={()=>excluir(confirmExcluir)} className="px-5 py-2 rounded-lg text-sm font-semibold text-white" style={{ background:"#c0392b" }}>Excluir</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ============================================================
// DASHBOARD GERENCIAL
// ============================================================
function DashboardGerencial({ notas, empresas }) {
  const hoje = new Date();
  const [dtDe, setDtDe] = useState(`01/${String(hoje.getMonth()+1).padStart(2,"0")}/${hoje.getFullYear()}`);
  const [dtAte, setDtAte] = useState(`${String(hoje.getDate()).padStart(2,"0")}/${String(hoje.getMonth()+1).padStart(2,"0")}/${hoje.getFullYear()}`);
  const [filtroEmp, setFiltroEmp] = useState("");
  const [filtroFin, setFiltroFin] = useState("");
  const [aplicado, setAplicado] = useState({ dtDe: `01/${String(hoje.getMonth()+1).padStart(2,"0")}/${hoje.getFullYear()}`, dtAte: `${String(hoje.getDate()).padStart(2,"0")}/${String(hoje.getMonth()+1).padStart(2,"0")}/${hoje.getFullYear()}`, emp: "", fin: "" });
  const buscarGerencial = () => setAplicado({ dtDe, dtAte, emp: filtroEmp, fin: filtroFin });
  const limparGerencial = () => { const d=`01/${String(hoje.getMonth()+1).padStart(2,"0")}/${hoje.getFullYear()}`; const a=`${String(hoje.getDate()).padStart(2,"0")}/${String(hoje.getMonth()+1).padStart(2,"0")}/${hoje.getFullYear()}`; setDtDe(d);setDtAte(a);setFiltroEmp("");setFiltroFin("");setAplicado({dtDe:d,dtAte:a,emp:"",fin:""}); };

  function parseBR(d) {
    if(!d||d.length<10) return null;
    const p=d.split("/"); if(p.length!==3) return null;
    return new Date(parseInt(p[2]),parseInt(p[1])-1,parseInt(p[0]));
  }
  const deA = parseBR(aplicado.dtDe), ateA = parseBR(aplicado.dtAte);
  const notasPeriodo = notas.filter(n => {
    const dt = parseBR(n.dtImportacao);
    if(!dt) return false;
    if(deA && dt < deA) return false;
    if(ateA && dt > ateA) return false;
    if(aplicado.emp && n.empresa !== aplicado.emp) return false;
    if(aplicado.fin && n.finalidade !== aplicado.fin) return false;
    return true;
  });

  const STATUS_FINAL = ["Desembaraçada","Selada","Recusada"];
  const resolvidas = notasPeriodo.filter(n => STATUS_FINAL.includes(n.status));
  const fmtM = v => v.toLocaleString("pt-BR",{style:"currency",currency:"BRL"});

  // Custos
  const totalCustos = notasPeriodo.reduce((s,n) => {
    const lancs = n.lancamentos||[];
    return s + lancs.reduce((a,l)=>a+(parseFloat(l.valor)||0),0);
  },0);
  const totalPago = notasPeriodo.reduce((s,n) => {
    const lancs = (n.lancamentos||[]).filter(l=>l.pago);
    return s + lancs.reduce((a,l)=>a+(parseFloat(l.valor)||0),0);
  },0);

  // Custos por tipo
  const custosPorTipo = {};
  notasPeriodo.forEach(n => (n.lancamentos||[]).forEach(l => {
    custosPorTipo[l.tipo] = (custosPorTipo[l.tipo]||0)+(parseFloat(l.valor)||0);
  }));
  const tiposSorted = Object.entries(custosPorTipo).sort((a,b)=>b[1]-a[1]);
  const maxTipo = tiposSorted[0]?.[1]||1;

  // Por empresa
  const empData = (empresas||[]).map(e => ({
    nome: e.nome, id: e.id,
    count: resolvidas.filter(n=>n.empresa===e.id).length,
    custo: notasPeriodo.filter(n=>n.empresa===e.id).reduce((s,n)=>s+(n.lancamentos||[]).reduce((a,l)=>a+(parseFloat(l.valor)||0),0),0)
  })).sort((a,b)=>b.count-a.count);
  const maxEmp = empData[0]?.count||1;

  // Por finalidade
  const finData = {};
  notasPeriodo.forEach(n=>{ const f=n.finalidade||"Não identificado"; finData[f]=(finData[f]||0)+1; });
  const finSorted = Object.entries(finData).sort((a,b)=>b[1]-a[1]);
  const maxFin = finSorted[0]?.[1]||1;

  // Por CC
  const ccData = {};
  notasPeriodo.forEach(n=>{ const c=n.centroCusto||"Sem CC"; ccData[c]=(ccData[c]||0)+(n.lancamentos||[]).reduce((a,l)=>a+(parseFloat(l.valor)||0),0); });
  const ccSorted = Object.entries(ccData).sort((a,b)=>b[1]-a[1]).slice(0,5);
  const maxCC = ccSorted[0]?.[1]||1;

  // Tempo médio por mês (últimos 6 meses)
  function calcTempoMedio(mes, ano) {
    const noMes = notas.filter(n => {
      if(!STATUS_FINAL.includes(n.status)) return false;
      const hist = (n.historico||[]).find(h => STATUS_FINAL.some(s=>h.acao?.includes(s)));
      if(!hist) return false;
      const dtFim = parseBR(hist.data?.substring(0,10).split("/").length===3?hist.data.substring(0,10):null);
      if(!dtFim) return false;
      return dtFim.getMonth()===mes && dtFim.getFullYear()===ano;
    });
    if(!noMes.length) return null;
    const dias = noMes.map(n=>{
      const dtImp=parseBR(n.dtImportacao); if(!dtImp) return 0;
      const hist=(n.historico||[]).find(h=>STATUS_FINAL.some(s=>h.acao?.includes(s)));
      if(!hist) return 0;
      const p=hist.data?.substring(0,10); const dtFim=parseBR(p); if(!dtFim) return 0;
      return Math.max(0,Math.floor((dtFim-dtImp)/86400000));
    }).filter(d=>d>0);
    return dias.length ? Math.round(dias.reduce((a,b)=>a+b,0)/dias.length) : null;
  }

  const meses6 = Array.from({length:6},(_,i)=>{
    const d=new Date(hoje.getFullYear(),hoje.getMonth()-5+i,1);
    return { mes:d.getMonth(), ano:d.getFullYear(), label:d.toLocaleString("pt-BR",{month:"short"}).replace(".","") };
  });
  const tempoMeses = meses6.map(m=>({ ...m, media: calcTempoMedio(m.mes,m.ano) }));
  const mediaGeral = (()=>{
    const vals=resolvidas.map(n=>{
      const dtImp=parseBR(n.dtImportacao); if(!dtImp) return 0;
      const hist=(n.historico||[]).find(h=>STATUS_FINAL.some(s=>h.acao?.includes(s)));
      if(!hist) return 0;
      const p=hist.data?.substring(0,10); const dtFim=parseBR(p); if(!dtFim) return 0;
      return Math.max(0,Math.floor((dtFim-dtImp)/86400000));
    }).filter(d=>d>0);
    return vals.length?Math.round(vals.reduce((a,b)=>a+b,0)/vals.length):0;
  })();
  const maxTempo = Math.max(...tempoMeses.map(m=>m.media||0),1);

  const BAR = ({pct,cor,label}) => (
    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:7}}>
      <div style={{width:95,textAlign:"right",fontSize:11,color:"#6b7280",flexShrink:0}}>{label}</div>
      <div style={{flex:1,height:20,background:"#f0f4f4",borderRadius:5,overflow:"hidden"}}>
        <div style={{width:`${pct}%`,height:"100%",background:"#1a4a4a",borderRadius:5,display:"flex",alignItems:"center",justifyContent:"flex-end",paddingRight:6}}>
          <span style={{fontSize:10,color:"#fff",fontWeight:500}}></span>
        </div>
      </div>
    </div>
  );

  return (
    <div className="space-y-5">
      {/* Header filtros */}
      <div className="rounded-2xl p-5" style={{ background:"#1a4a4a" }}>
        <div className="flex items-center justify-between flex-wrap gap-3">
          <div>
            <h2 className="font-bold text-white text-base">Dashboard Gerencial — DTE/AM</h2>
            <p className="text-xs mt-1" style={{ color:"#7ecece" }}>Enerwatt Engenharia · Departamento Fiscal</p>
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            <input value={dtDe} onChange={e=>setDtDe(e.target.value)} placeholder="DD/MM/AAAA"
              className="border rounded-lg px-3 py-1.5 text-sm" style={{ borderColor:"rgba(126,206,206,0.3)",background:"rgba(255,255,255,0.08)",color:"#fff",width:110 }} />
            <span className="text-xs" style={{ color:"#7ecece" }}>até</span>
            <input value={dtAte} onChange={e=>setDtAte(e.target.value)} placeholder="DD/MM/AAAA"
              className="border rounded-lg px-3 py-1.5 text-sm" style={{ borderColor:"rgba(126,206,206,0.3)",background:"rgba(255,255,255,0.08)",color:"#fff",width:110 }} />
            <select value={filtroEmp} onChange={e=>setFiltroEmp(e.target.value)}
              className="border rounded-lg px-3 py-1.5 text-sm" style={{ borderColor:"rgba(126,206,206,0.3)",background:"rgba(255,255,255,0.08)",color:"#fff" }}>
              <option value="">Todas as empresas</option>
              {(empresas||[]).map(e=><option key={e.id} value={e.id} style={{background:"#1a4a4a"}}>{e.nome}</option>)}
            </select>
            <select value={filtroFin} onChange={e=>setFiltroFin(e.target.value)}
              className="border rounded-lg px-3 py-1.5 text-sm" style={{ borderColor:"rgba(126,206,206,0.3)",background:"rgba(255,255,255,0.08)",color:"#fff" }}>
              <option value="">Todas as finalidades</option>
              {["Uso/Consumo","Industrialização","Revenda","Remessa/Transferência","Imobilizado","Não Identificado"].map(f=><option key={f} value={f} style={{background:"#1a4a4a"}}>{f}</option>)}
            </select>
            <button onClick={buscarGerencial}
              className="px-5 py-1.5 rounded-lg text-sm font-bold flex items-center gap-2"
              style={{ background:"#E8450A", color:"#fff", border:"none" }}>
              🔍 Buscar
            </button>
            <button onClick={limparGerencial}
              className="px-4 py-1.5 rounded-lg text-sm font-medium"
              style={{ border:"1px solid rgba(126,206,206,0.3)", background:"transparent", color:"#7ecece" }}>
              Limpar
            </button>
          </div>
        </div>
      </div>

      {/* KPIs */}
      <div className="grid gap-3" style={{ gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))" }}>
        {[
          { label:"Notas no período", val:notasPeriodo.length, sub:"importadas", cor:"#1a4a4a", bg:"#edf5f5", acc:"#1a4a4a" },
          { label:"Resolvidas", val:resolvidas.length, sub:`${notasPeriodo.length?Math.round(resolvidas.length/notasPeriodo.length*100):0}% do total`, cor:"#2d6a4f", bg:"#f0fdf4", acc:"#2d6a4f" },
          { label:"Custo total", val:fmtM(totalCustos), sub:`Pago: ${fmtM(totalPago)}`, cor:"#E8450A", bg:"#fff8f0", acc:"#E8450A" },
          { label:"Pendente pgto", val:fmtM(totalCustos-totalPago), sub:"a pagar", cor:"#b7791f", bg:"#fffbeb", acc:"#b7791f" },
          { label:"Tempo médio", val:mediaGeral?`${mediaGeral} dias`:"—", sub:"para resolver", cor:"#4db8b8", bg:"#edf5f5", acc:"#4db8b8" },
        ].map((k,i)=>(
          <div key={i} className="rounded-2xl p-4 relative overflow-hidden" style={{ background:k.bg, border:`1px solid ${k.cor}22` }}>
            <div style={{ position:"absolute",top:0,left:0,width:4,height:"100%",background:k.acc,borderRadius:"12px 0 0 12px" }} />
            <p className="text-xs font-semibold uppercase tracking-wide" style={{ color:"#6b7280" }}>{k.label}</p>
            <p className="text-xl font-black mt-1" style={{ color:k.cor }}>{k.val}</p>
            <p className="text-xs mt-1" style={{ color:"#9ca3af" }}>{k.sub}</p>
          </div>
        ))}
      </div>

      {/* Gráficos linha 1 */}
      <div className="grid grid-cols-2 gap-4">
        {/* Composição custos */}
        <div className="rounded-2xl border p-5" style={{ borderColor:"#f0f0f0" }}>
          <h3 className="font-bold text-gray-700 text-sm mb-4">Composição dos custos</h3>
          {tiposSorted.length === 0 ? <p className="text-xs text-gray-400">Nenhum lançamento no período</p> :
            tiposSorted.map(([tipo,val],i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                <div style={{width:110,textAlign:"right",fontSize:11,color:"#6b7280",flexShrink:0}}>{tipo}</div>
                <div style={{flex:1,height:20,background:"#f0f4f4",borderRadius:5,overflow:"hidden"}}>
                  <div style={{width:`${Math.round(val/maxTipo*100)}%`,height:"100%",background:i===0?"#1a4a4a":i===1?"#E8450A":"#4db8b8",borderRadius:5,display:"flex",alignItems:"center",justifyContent:"flex-end",paddingRight:6}}>
                    <span style={{fontSize:10,color:"#fff",fontWeight:500}}>{fmtM(val)}</span>
                  </div>
                </div>
              </div>
            ))
          }
        </div>

        {/* Por empresa */}
        <div className="rounded-2xl border p-5" style={{ borderColor:"#f0f0f0" }}>
          <h3 className="font-bold text-gray-700 text-sm mb-4">Resolvidas por empresa</h3>
          {empData.map((e,i)=>(
            <div key={i} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
              <div style={{width:110,textAlign:"right",fontSize:11,color:"#6b7280",flexShrink:0}}>{e.nome.replace("FILIAL","Filial").replace("LINHAS DO NORTE","Linhas do Norte")}</div>
              <div style={{flex:1,height:20,background:"#f0f4f4",borderRadius:5,overflow:"hidden"}}>
                <div style={{width:`${e.count?Math.round(e.count/maxEmp*100):0}%`,height:"100%",background:i===0?"#1a4a4a":"#4db8b8",borderRadius:5,display:"flex",alignItems:"center",justifyContent:"flex-end",paddingRight:6}}>
                  <span style={{fontSize:10,color:"#fff",fontWeight:500}}>{e.count}</span>
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Gráficos linha 2 */}
      <div className="grid grid-cols-2 gap-4">
        {/* Por finalidade */}
        <div className="rounded-2xl border p-5" style={{ borderColor:"#f0f0f0" }}>
          <h3 className="font-bold text-gray-700 text-sm mb-4">Por finalidade</h3>
          {finSorted.length===0 ? <p className="text-xs text-gray-400">Sem dados no período</p> :
            finSorted.map(([fin,cnt],i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                <div style={{width:110,textAlign:"right",fontSize:11,color:"#6b7280",flexShrink:0}}>{fin}</div>
                <div style={{flex:1,height:20,background:"#f0f4f4",borderRadius:5,overflow:"hidden"}}>
                  <div style={{width:`${Math.round(cnt/maxFin*100)}%`,height:"100%",background:i===0?"#1a4a4a":i===1?"#E8450A":"#4db8b8",borderRadius:5,display:"flex",alignItems:"center",justifyContent:"flex-end",paddingRight:6}}>
                    <span style={{fontSize:10,color:"#fff",fontWeight:500}}>{cnt}</span>
                  </div>
                </div>
              </div>
            ))
          }
        </div>

        {/* Por CC */}
        <div className="rounded-2xl border p-5" style={{ borderColor:"#f0f0f0" }}>
          <h3 className="font-bold text-gray-700 text-sm mb-4">Custo por centro de custo</h3>
          {ccSorted.length===0 ? <p className="text-xs text-gray-400">Sem dados no período</p> :
            ccSorted.map(([cc,val],i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                <div style={{width:110,textAlign:"right",fontSize:11,color:"#6b7280",flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{cc}</div>
                <div style={{flex:1,height:20,background:"#f0f4f4",borderRadius:5,overflow:"hidden"}}>
                  <div style={{width:`${Math.round(val/maxCC*100)}%`,height:"100%",background:i===0?"#1a4a4a":"#7ecece",borderRadius:5,display:"flex",alignItems:"center",justifyContent:"flex-end",paddingRight:6}}>
                    <span style={{fontSize:10,color:i===0?"#fff":"#1a4a4a",fontWeight:500}}>{fmtM(val)}</span>
                  </div>
                </div>
              </div>
            ))
          }
        </div>
      </div>

      {/* Tempo médio de resolução */}
      <div className="rounded-2xl border p-5" style={{ borderColor:"#f0f0f0" }}>
        <h3 className="font-bold text-gray-700 text-sm mb-1">Tempo médio de resolução</h3>
        <p className="text-xs text-gray-400 mb-4">Dias da importação até Desembaraçada, Selada ou Recusada — por mês</p>
        <div style={{ display:"flex",alignItems:"flex-end",gap:6,height:100 }}>
          {tempoMeses.map((m,i)=>{
            const h = m.media ? Math.round((m.media/maxTempo)*85) : 0;
            const cor = !m.media?"#e5e7eb":m.media<=15?"#2d6a4f":m.media<=30?"#b7791f":"#c0392b";
            const isAtual = i===tempoMeses.length-1;
            return (
              <div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
                <span style={{fontSize:10,fontWeight:500,color:m.media?cor:"#d1d5db"}}>{m.media?`${m.media}d`:"—"}</span>
                <div style={{width:"100%",height:h||4,background:isAtual?"#1a4a4a":cor,borderRadius:"4px 4px 0 0",opacity:isAtual?1:0.7}}/>
                <span style={{fontSize:9,color:"#9ca3af"}}>{m.label}</span>
              </div>
            );
          })}
        </div>
        <div className="mt-4 flex items-center gap-3 pt-3" style={{ borderTop:"1px solid #f0f0f0" }}>
          <span className="text-xs font-bold px-4 py-1.5 rounded-full text-white" style={{ background:"#1a4a4a" }}>
            Média geral: {mediaGeral||"—"} {mediaGeral?"dias":""}
          </span>
          <span className="text-xs text-gray-400">Importação → status final (Desembaraçada, Selada ou Recusada)</span>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// LOG DE ACESSOS
// ============================================================
function LogAcessoBtn() {
  const [aberto, setAberto] = useState(false);
  const [logs, setLogs] = useState([]);
  const [carregando, setCarregando] = useState(false);

  async function carregar() {
    setCarregando(true);
    const { data } = await supabase
      .from("logs_acesso")
      .select("*")
      .order("acessado_em", { ascending: false })
      .limit(50);
    if (data) setLogs(data);
    setCarregando(false);
  }

  function abrir() { setAberto(true); carregar(); }

  const fmtDt = (iso) => {
    if (!iso) return "—";
    return new Date(iso).toLocaleString("pt-BR", {
      day: "2-digit", month: "2-digit", year: "numeric",
      hour: "2-digit", minute: "2-digit"
    });
  };

  return (
    <>
      <button onClick={abrir}
        className="text-xs px-3 py-1.5 rounded-lg font-semibold flex items-center gap-1.5"
        style={{ background:"#edf5f5", color:"#1a4a4a", border:"1px solid #d0e8e8" }}>
        📋 Histórico de Acessos
      </button>

      {aberto && (
        <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background:"rgba(0,0,0,0.5)" }}>
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg m-4 flex flex-col" style={{ maxHeight:"80vh" }}>
            <div className="flex items-center justify-between p-5 border-b" style={{ borderColor:"#f0f0f0" }}>
              <div>
                <h2 className="font-bold text-gray-800">Histórico de Acessos</h2>
                <p className="text-xs text-gray-400 mt-0.5">Últimos 50 logins registrados</p>
              </div>
              <div className="flex items-center gap-2">
                <button onClick={carregar} className="text-xs px-3 py-1.5 rounded-lg border font-medium" style={{ borderColor:"#e5e7eb", color:"#374151" }}>
                  🔄 Atualizar
                </button>
                <button onClick={() => setAberto(false)} className="text-gray-400 hover:text-gray-600 text-2xl font-light leading-none">×</button>
              </div>
            </div>

            <div className="overflow-y-auto flex-1 p-4 space-y-2">
              {carregando && (
                <div className="text-center py-8 text-gray-400 text-sm">Carregando...</div>
              )}
              {!carregando && logs.length === 0 && (
                <div className="text-center py-8 text-gray-400 text-sm">Nenhum acesso registrado ainda.</div>
              )}
              {logs.map((l, i) => (
                <div key={i} className="flex items-center gap-3 p-3 rounded-xl" style={{ background:"#f8f9fa" }}>
                  <div className="w-8 h-8 rounded-full flex items-center justify-center text-white text-sm font-bold flex-shrink-0" style={{ background:"#1a4a4a" }}>
                    {(l.usuario_nome||"?").charAt(0).toUpperCase()}
                  </div>
                  <div className="flex-1 min-w-0">
                    <p className="font-semibold text-sm text-gray-800">{l.usuario_nome}</p>
                    <p className="text-xs text-gray-400">{l.usuario_email}</p>
                  </div>
                  <div className="text-right flex-shrink-0">
                    <p className="text-xs font-semibold" style={{ color:"#1a4a4a" }}>{fmtDt(l.acessado_em)}</p>
                    <span className={`text-xs px-2 py-0.5 rounded-full font-semibold ${l.perfil === "admin" ? "bg-orange-100 text-orange-700" : "bg-blue-100 text-blue-700"}`}>
                      {l.perfil}
                    </span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </>
  );
}

// ============================================================
// MODAL ALTERAR SENHA
// ============================================================
function ModalAlterarSenha({ onClose }) {
  const [senhaAtual, setSenhaAtual] = useState("");
  const [novaSenha, setNovaSenha] = useState("");
  const [confirmar, setConfirmar] = useState("");
  const [loading, setLoading] = useState(false);
  const [msg, setMsg] = useState(null); // { tipo: "ok"|"erro", texto }

  const salvar = async () => {
    if (novaSenha.length < 6) { setMsg({ tipo: "erro", texto: "A nova senha deve ter ao menos 6 caracteres." }); return; }
    if (novaSenha !== confirmar) { setMsg({ tipo: "erro", texto: "As senhas não coincidem." }); return; }
    setLoading(true);
    setMsg(null);
    try {
      // Reautentica com senha atual para validar
      const { data: { user } } = await supabase.auth.getUser();
      const { error: reAuthError } = await supabase.auth.signInWithPassword({ email: user.email, password: senhaAtual });
      if (reAuthError) { setMsg({ tipo: "erro", texto: "Senha atual incorreta." }); setLoading(false); return; }
      // Atualiza senha
      const { error } = await supabase.auth.updateUser({ password: novaSenha });
      if (error) { setMsg({ tipo: "erro", texto: "Erro ao alterar senha: " + error.message }); setLoading(false); return; }
      setMsg({ tipo: "ok", texto: "✅ Senha alterada com sucesso!" });
      setTimeout(() => onClose(), 1800);
    } catch (e) {
      setMsg({ tipo: "erro", texto: "Erro inesperado: " + e.message });
    }
    setLoading(false);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background: "rgba(0,0,0,0.5)" }}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-sm m-4">
        <div className="flex items-center justify-between p-5 border-b" style={{ borderColor: "#f0f0f0" }}>
          <div>
            <h2 className="font-bold text-gray-800">🔒 Alterar Senha</h2>
            <p className="text-xs text-gray-400 mt-0.5">Defina uma nova senha para sua conta</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 text-2xl font-light">×</button>
        </div>
        <div className="p-5 space-y-4">
          <div>
            <label className="text-xs font-semibold text-gray-500 uppercase">Senha Atual</label>
            <input type="password" value={senhaAtual} onChange={e => setSenhaAtual(e.target.value)}
              placeholder="Digite sua senha atual"
              className="mt-1 w-full border rounded-lg p-2.5 text-sm" style={{ borderColor: "#e5e7eb" }} />
          </div>
          <div>
            <label className="text-xs font-semibold text-gray-500 uppercase">Nova Senha</label>
            <input type="password" value={novaSenha} onChange={e => setNovaSenha(e.target.value)}
              placeholder="Mínimo 6 caracteres"
              className="mt-1 w-full border rounded-lg p-2.5 text-sm" style={{ borderColor: "#e5e7eb" }} />
          </div>
          <div>
            <label className="text-xs font-semibold text-gray-500 uppercase">Confirmar Nova Senha</label>
            <input type="password" value={confirmar} onChange={e => setConfirmar(e.target.value)}
              placeholder="Repita a nova senha"
              className="mt-1 w-full border rounded-lg p-2.5 text-sm" style={{ borderColor: "#e5e7eb" }} />
          </div>
          {msg && (
            <div className="p-3 rounded-lg text-sm font-medium" style={{
              background: msg.tipo === "ok" ? "#f0fdf4" : "#fff0f0",
              color: msg.tipo === "ok" ? "#2d6a4f" : "#c0392b",
              border: `1px solid ${msg.tipo === "ok" ? "#d1fae5" : "#ffc7c7"}`
            }}>
              {msg.texto}
            </div>
          )}
        </div>
        <div className="flex justify-end gap-3 p-5 border-t" style={{ borderColor: "#f0f0f0" }}>
          <button onClick={onClose} className="px-4 py-2 rounded-lg border text-sm text-gray-600" style={{ borderColor: "#e5e7eb" }}>Cancelar</button>
          <button onClick={salvar} disabled={loading || !senhaAtual || !novaSenha || !confirmar}
            className="px-6 py-2 rounded-lg text-sm font-semibold text-white disabled:opacity-50"
            style={{ background: "#1a4a4a" }}>
            {loading ? "Salvando..." : "Salvar Nova Senha"}
          </button>
        </div>
      </div>
    </div>
  );
}

function Configuracoes({ usuarios, onSalvarUsuario, onEditarUsuario, onExcluirUsuario, logoUrl, onSalvarLogo, empresas, onSalvarEmpresa, perfilAtual }) {
  const [novoNome, setNovoNome] = useState("");
  const [novoEmail, setNovoEmail] = useState("");
  const [novoPerfil, setNovoPerfil] = useState("operador");
  const [modalEmpresa, setModalEmpresa] = useState(null);
  const [modalUsuario, setModalUsuario] = useState(null); // null = fechado, {id,...} = editar
  const [modalSenha, setModalSenha] = useState(false);

  const handleLogo = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => onSalvarLogo(ev.target.result);
    reader.readAsDataURL(file);
  };

  return (
    <div className="space-y-6">

      {/* Minha Conta */}
      <div className="rounded-2xl border p-5" style={{ borderColor: "#f0f0f0" }}>
        <h3 className="font-bold text-gray-700 mb-1 text-sm uppercase tracking-wide">Minha Conta</h3>
        <p className="text-xs text-gray-400 mb-4">Gerencie as credenciais de acesso da sua conta.</p>
        <button onClick={() => setModalSenha(true)}
          className="px-5 py-2.5 rounded-xl text-sm font-semibold text-white hover:opacity-90 transition-opacity"
          style={{ background: "#1a4a4a" }}>
          🔒 Alterar Minha Senha
        </button>
        <p className="text-xs text-gray-400 mt-3">
          Caso esqueça sua senha, peça ao administrador do sistema para redefini-la.
        </p>
      </div>

      {modalSenha && <ModalAlterarSenha onClose={() => setModalSenha(false)} />}

      {/* Logo */}
      <div className="rounded-2xl border p-5" style={{ borderColor: "#f0f0f0" }}>
        <h3 className="font-bold text-gray-700 mb-4 text-sm uppercase tracking-wide">Logo do Sistema</h3>
        <p className="text-xs text-gray-400 mb-4">Aparece na sidebar e será incluída nos relatórios exportados.</p>
        <div className="flex items-center gap-5">
          <div className="w-40 h-20 rounded-xl border flex items-center justify-center overflow-hidden" style={{ borderColor: "#e5e7eb", background: "#1a1a1a" }}>
            {logoUrl ? <img src={logoUrl} alt="Logo" className="max-w-full max-h-full object-contain p-2" />
              : <span className="text-xs text-gray-500">Sem logo</span>}
          </div>
          <div>
            <label className="px-4 py-2 rounded-lg text-sm font-semibold text-white cursor-pointer inline-block" style={{ background: "#E8450A" }}>
              📷 Alterar Logo
              <input type="file" accept="image/*" className="hidden" onChange={handleLogo} />
            </label>
            {logoUrl && <button onClick={() => onSalvarLogo(null)} className="ml-3 text-xs text-gray-400 hover:text-red-500">Remover</button>}
            <p className="text-xs text-gray-400 mt-2">PNG, JPG ou SVG. Fundo transparente ou escuro recomendado.</p>
          </div>
        </div>
      </div>

      {/* Empresas */}
      <div className="rounded-2xl border p-5" style={{ borderColor: "#f0f0f0" }}>
        <div className="flex items-center justify-between mb-4">
          <h3 className="font-bold text-gray-700 text-sm uppercase tracking-wide">Empresas / Inscrições Estaduais</h3>
          <button onClick={() => setModalEmpresa({})}
            className="px-4 py-2 rounded-lg text-sm font-semibold text-white" style={{ background: "#E8450A" }}>
            + Nova Empresa
          </button>
        </div>
        <p className="text-xs text-gray-400 mb-4">O CNPJ e a IE são usados para identificar automaticamente a empresa durante o upload do CSV.</p>
        <div className="space-y-3">
          {empresas.map(e => (
            <div key={e.id} className="flex items-center justify-between p-4 rounded-xl" style={{ background: "#f8f9fa", border: "1px solid #f0f0f0" }}>
              <div className="flex-1">
                <div className="flex items-center gap-2">
                  <p className="font-semibold text-sm text-gray-800">{e.nome}</p>
                  <span className={`text-xs px-2 py-0.5 rounded-full font-semibold ${e.ativa ? "bg-green-100 text-green-700" : "bg-gray-100 text-gray-500"}`}>
                    {e.ativa ? "Ativa" : "Inativa"}
                  </span>
                </div>
                <div className="flex gap-4 mt-1">
                  <p className="text-xs text-gray-400">IE: <span className="font-mono text-gray-600">{e.inscricao}</span></p>
                  <p className="text-xs text-gray-400">CNPJ: <span className="font-mono text-gray-600">{e.cnpj}</span></p>
                </div>
              </div>
              <button onClick={() => setModalEmpresa(e)}
                className="px-3 py-1.5 rounded-lg text-xs font-semibold border ml-3" style={{ borderColor: "#e5e7eb", color: "#374151" }}>
                ✏️ Editar
              </button>
            </div>
          ))}
        </div>
      </div>

      {/* Usuários */}
      <div className="rounded-2xl border p-5" style={{ borderColor: "#f0f0f0" }}>
        <div className="flex items-center justify-between mb-4">
          <h3 className="font-bold text-gray-700 text-sm uppercase tracking-wide">Usuários</h3>
          {perfilAtual === "admin" && (
            <LogAcessoBtn />
          )}
        </div>
        <div className="space-y-2 mb-5">
          {usuarios.map(u => {
            const ultimoAcesso = u.ultimo_acesso
              ? new Date(u.ultimo_acesso).toLocaleString("pt-BR", { day:"2-digit", month:"2-digit", year:"numeric", hour:"2-digit", minute:"2-digit" })
              : null;
            return (
            <div key={u.id} className="flex items-center justify-between p-3 rounded-xl" style={{ background: "#f8f9fa" }}>
              <div className="flex items-center gap-3">
                <div className="w-8 h-8 rounded-full flex items-center justify-center text-white text-sm font-bold" style={{ background: "#E8450A" }}>{u.nome.charAt(0)}</div>
                <div>
                  <p className="font-semibold text-sm text-gray-800">{u.nome}</p>
                  <p className="text-xs text-gray-400">{u.email}</p>
                  {ultimoAcesso && (
                    <p className="text-xs mt-0.5" style={{ color:"#4db8b8" }}>
                      Último acesso: {ultimoAcesso}
                    </p>
                  )}
                  {!ultimoAcesso && (
                    <p className="text-xs mt-0.5 text-gray-300">Ainda não acessou</p>
                  )}
                </div>
              </div>
              <div className="flex items-center gap-2">
                <span className={`text-xs px-2 py-1 rounded-full font-semibold ${u.perfil === "admin" ? "bg-orange-100 text-orange-700" : "bg-blue-100 text-blue-700"}`}>{u.perfil}</span>
                <button onClick={() => setModalUsuario(u)} className="text-xs px-2 py-1 rounded-lg border font-semibold" style={{ borderColor: "#e5e7eb", color: "#374151" }}>✏️</button>
                {perfilAtual === "admin" && (
                  <button onClick={() => { if (window.confirm(`Excluir usuário ${u.nome}?`)) onExcluirUsuario(u.id); }}
                    className="text-xs px-2 py-1 rounded-lg border font-semibold" style={{ borderColor: "#fecaca", color: "#dc2626" }}>🗑️</button>
                )}
              </div>
            </div>
            );
          })}
        </div>
        <h4 className="text-sm font-semibold text-gray-600 mb-3">Adicionar Usuário</h4>
        <p className="text-xs text-gray-400 mb-3">A senha inicial será <strong>Enerwatt@2024</strong> — o usuário pode alterar depois.</p>
        <div className="grid grid-cols-2 gap-3">
          <input placeholder="Nome completo" value={novoNome} onChange={e => setNovoNome(e.target.value)} className="border rounded-lg px-3 py-2 text-sm col-span-2" style={{ borderColor: "#e5e7eb" }} />
          <input placeholder="E-mail" value={novoEmail} onChange={e => setNovoEmail(e.target.value)} className="border rounded-lg px-3 py-2 text-sm" style={{ borderColor: "#e5e7eb" }} />
          <select value={novoPerfil} onChange={e => setNovoPerfil(e.target.value)} className="border rounded-lg px-3 py-2 text-sm" style={{ borderColor: "#e5e7eb" }}>
            <option value="operador">Operador</option>
            <option value="admin">Admin</option>
          </select>
          <button onClick={() => { if (novoNome && novoEmail) { onSalvarUsuario({ id: Date.now().toString(), nome: novoNome, email: novoEmail, perfil: novoPerfil }); setNovoNome(""); setNovoEmail(""); setNovoPerfil("operador"); } }}
            className="col-span-2 py-2 rounded-lg text-sm font-semibold text-white" style={{ background: "#E8450A" }}>
            Adicionar Usuário
          </button>
        </div>
      </div>

      {modalEmpresa !== null && (
        <ModalEmpresa
          empresa={modalEmpresa.id ? modalEmpresa : null}
          onClose={() => setModalEmpresa(null)}
          onSalvar={(e) => { onSalvarEmpresa(e); setModalEmpresa(null); }}
        />
      )}

      {modalUsuario !== null && (
        <ModalEditarUsuario
          usuario={modalUsuario}
          onClose={() => setModalUsuario(null)}
          onSalvar={(u) => { onEditarUsuario(u); setModalUsuario(null); }}
          perfilAtual={perfilAtual}
        />
      )}
    </div>
  );
}

// ============================================================
// MODAL EDITAR USUARIO
// ============================================================
function ModalEditarUsuario({ usuario, onClose, onSalvar, perfilAtual }) {
  const [form, setForm] = useState({ ...usuario });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background: "rgba(0,0,0,0.5)" }}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md m-4">
        <div className="flex items-center justify-between p-5 border-b" style={{ borderColor: "#f0f0f0" }}>
          <h2 className="font-bold text-gray-800">Editar Usuário</h2>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 text-2xl font-light">×</button>
        </div>
        <div className="p-5 space-y-4">
          <div>
            <label className="text-xs font-semibold text-gray-500 uppercase">Nome completo</label>
            <input value={form.nome} onChange={e => set("nome", e.target.value)}
              className="mt-1 w-full border rounded-lg px-3 py-2 text-sm" style={{ borderColor: "#e5e7eb" }} />
          </div>
          <div>
            <label className="text-xs font-semibold text-gray-500 uppercase">E-mail</label>
            <input value={form.email} onChange={e => set("email", e.target.value)}
              className="mt-1 w-full border rounded-lg px-3 py-2 text-sm" style={{ borderColor: "#e5e7eb" }} />
          </div>
          {perfilAtual === "admin" && (
            <div>
              <label className="text-xs font-semibold text-gray-500 uppercase">Perfil</label>
              <select value={form.perfil} onChange={e => set("perfil", e.target.value)}
                className="mt-1 w-full border rounded-lg px-3 py-2 text-sm" style={{ borderColor: "#e5e7eb" }}>
                <option value="operador">Operador</option>
                <option value="admin">Admin</option>
              </select>
            </div>
          )}
        </div>
        <div className="flex justify-end gap-3 p-5 border-t" style={{ borderColor: "#f0f0f0" }}>
          <button onClick={onClose} className="px-4 py-2 rounded-lg border text-sm text-gray-600" style={{ borderColor: "#e5e7eb" }}>Cancelar</button>
          <button onClick={() => onSalvar(form)} className="px-6 py-2 rounded-lg text-sm font-semibold text-white" style={{ background: "#E8450A" }}>
            Salvar Alterações
          </button>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// APP PRINCIPAL
// ============================================================
export default function App() {
  const [tela, setTela] = useState("dashboard");
  const [notas, setNotas] = useState([]);
  const [notaSelecionada, setNotaSelecionada] = useState(null);
  const [usuarios, setUsuarios] = useState([]);
  const [usuarioAtual, setUsuarioAtual] = useState(null);
  const [ultimaImportacao, setUltimaImportacao] = useState("-");
  const [logoUrl, setLogoUrl] = useState(null);
  const [empresas, setEmpresas] = useState(EMPRESAS_INICIAIS);
  const [importPendente, setImportPendente] = useState(null);
  const [carregando, setCarregando] = useState(true);
  const [loginCarregando, setLoginCarregando] = useState(true);

  // Atualiza alias global
  EMPRESAS = empresas;

  // Verificar sessão ativa ao carregar
  useEffect(() => {
    const verificarSessao = async () => {
      try {
        const { data: { session } } = await supabase.auth.getSession();
        if (session) {
          try {
            const { data } = await supabase.from("usuarios").select("*").eq("email", session.user.email).single();
            setUsuarioAtual({ ...session.user, nome: data?.nome || session.user.email, perfil: data?.perfil || "operador" });
          } catch(e) {
            setUsuarioAtual({ ...session.user, nome: session.user.email, perfil: "operador" });
          }
        }
      } catch(e) {
        console.error("Erro ao verificar sessao:", e);
      } finally {
        setLoginCarregando(false);
      }
    };
    verificarSessao();
    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => {
      if (!session) setUsuarioAtual(null);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  // ---- CARREGAR DADOS DO SUPABASE ----
  useEffect(() => {
    if (!usuarioAtual) return;
    carregarTudo();
    // Realtime: atualiza automaticamente quando outro usuário mudar dados
    const canal = supabase
      .channel("realtime-notas")
      .on("postgres_changes", { event: "*", schema: "public", table: "notas" }, () => carregarNotas())
      .on("postgres_changes", { event: "*", schema: "public", table: "empresas" }, () => carregarEmpresas())
      .subscribe();
    return () => supabase.removeChannel(canal);
  }, [usuarioAtual]);

  async function carregarTudo() {
    setCarregando(true);
    await Promise.all([carregarNotas(), carregarEmpresas(), carregarConfig(), carregarUsuarios()]);
    setCarregando(false);
  }

  async function carregarUsuarios() {
    const { data } = await supabase.from("usuarios").select("*").order("nome");
    if (data) setUsuarios(data);
  }

  async function carregarNotas() {
    const { data } = await supabase.from("notas").select("*").order("criado_em", { ascending: false });
    if (data) setNotas(data.map(dbParaNota));
  }

  async function carregarEmpresas() {
    const { data } = await supabase.from("empresas").select("*").order("nome");
    if (data && data.length > 0) {
      setEmpresas(data);
      EMPRESAS = data;
    }
  }

  async function carregarConfig() {
    const { data } = await supabase.from("configuracoes").select("*").eq("id", "global").single();
    if (data) {
      if (data.logo_url) setLogoUrl(data.logo_url);
      if (data.ultima_importacao) setUltimaImportacao(data.ultima_importacao);
    }
  }

  // Converte registro do banco para formato do app
  function calcDiasEmissao(dtEmissao) {
    if (!dtEmissao || dtEmissao === "-") return 0;
    try {
      const parts = dtEmissao.split("/");
      if (parts.length !== 3) return 0;
      const dt = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
      const hoje = new Date(); hoje.setHours(0,0,0,0);
      return Math.max(0, Math.floor((hoje - dt) / 86400000));
    } catch { return 0; }
  }

  function dbParaNota(r) {
    return {
      id: r.id, empresa: r.empresa, fornecedor: r.fornecedor, razaoSocial: r.razao_social,
      numNota: r.num_nota, cfop: r.cfop, dtEmissao: r.dt_emissao, dtApresentacao: r.dt_apresentacao,
      chave: r.chave, valor: r.valor, qtdeDias: calcDiasEmissao(r.dt_emissao), status: r.status,
      centroCusto: r.centro_custo, codigoCusto: r.codigo_custo || "", finalidade: r.finalidade, responsavel: r.responsavel,
      lancamentos: r.lancamentos || [],
      noRM: r.no_rm, taxaReanalise: r.taxa_reanalise, taxaDesembaraco: r.taxa_desembaraco,
      icmsAntecipado: r.icms_antecipado, multa: r.multa, juros: r.juros,
      dtReanalise: r.dt_reanalise, dtDesembaraco: r.dt_desembaraco, dtPostergacao: r.dt_postergacao,
      multa10pct: r.multa10pct || 0, dtVencReanalise: r.dt_venc_reanalise || "", dtVencDesembaraco: r.dt_venc_desembaraco || "", dtVencIcms: r.dt_venc_icms || "",
      pagoReanalise: r.pago_reanalise || false, pagoDesembaraco: r.pago_desembaraco || false,
      pagoIcms: r.pago_icms || false, pagoMulta10: r.pago_multa10 || false,
      pagoMulta: r.pago_multa || false, pagoJuros: r.pago_juros || false,
      dtSelada: r.dt_selada || "",
      obs: r.obs, dtImportacao: r.dt_importacao, historico: r.historico || []
    };
  }

  // Converte formato do app para banco
  function notaParaDb(n) {
    return {
      id: n.id, empresa: n.empresa, fornecedor: n.fornecedor, razao_social: n.razaoSocial,
      num_nota: n.numNota, cfop: n.cfop, dt_emissao: n.dtEmissao, dt_apresentacao: n.dtApresentacao,
      chave: n.chave, valor: n.valor, qtde_dias: n.qtdeDias, status: n.status,
      centro_custo: n.centroCusto, codigo_custo: n.codigoCusto || "", finalidade: n.finalidade, responsavel: n.responsavel,
      lancamentos: n.lancamentos || [],
      no_rm: n.noRM, taxa_reanalise: n.taxaReanalise, taxa_desembaraco: n.taxaDesembaraco,
      icms_antecipado: n.icmsAntecipado, multa: n.multa, juros: n.juros,
      dt_reanalise: n.dtReanalise, dt_desembaraco: n.dtDesembaraco, dt_postergacao: n.dtPostergacao,
      multa10pct: n.multa10pct || 0, dt_venc_reanalise: n.dtVencReanalise || "", dt_venc_desembaraco: n.dtVencDesembaraco || "", dt_venc_icms: n.dtVencIcms || "",
      pago_reanalise: n.pagoReanalise || false, pago_desembaraco: n.pagoDesembaraco || false,
      pago_icms: n.pagoIcms || false, pago_multa10: n.pagoMulta10 || false,
      pago_multa: n.pagoMulta || false, pago_juros: n.pagoJuros || false,
      dt_selada: n.dtSelada || "",
      obs: n.obs, dt_importacao: n.dtImportacao, historico: n.historico,
      atualizado_em: new Date().toISOString()
    };
  }

  const handleSalvarNota = async (notaAtualizada) => {
    const { error } = await supabase.from("notas").upsert(notaParaDb(notaAtualizada));
    if (error) { alert("Erro ao salvar nota: " + error.message); return; }
    setNotas(ns => ns.map(n => n.id === notaAtualizada.id ? notaAtualizada : n));
    setNotaSelecionada(null);
  };

  const handleSalvarEmpresa = async (emp) => {
    const { error } = await supabase.from("empresas").upsert({ id: emp.id, nome: emp.nome, inscricao: emp.inscricao, cnpj: emp.cnpj, ativa: emp.ativa });
    if (error) { alert("Erro ao salvar empresa: " + error.message); return; }
    setEmpresas(es => {
      const idx = es.findIndex(e => e.id === emp.id);
      if (idx >= 0) return es.map(e => e.id === emp.id ? emp : e);
      return [...es, emp];
    });
  };

  const handleSalvarLogo = async (url) => {
    setLogoUrl(url);
    await supabase.from("configuracoes").upsert({ id: "global", logo_url: url, atualizado_em: new Date().toISOString() });
  };

  const handleSalvarUsuario = async (u) => {
    // Cria usuário no Supabase Auth
    const { data, error } = await supabase.auth.admin ? 
      { data: null, error: { message: "use tabela" } } :
      { data: null, error: { message: "use tabela" } };
    // Salva na tabela usuarios
    const { error: err } = await supabase.from("usuarios").upsert({ id: u.id, nome: u.nome, email: u.email, perfil: u.perfil });
    if (err) { alert("Erro ao salvar usuário: " + err.message); return; }
    await carregarUsuarios();
  };

  const handleEditarUsuario = async (u) => {
    const { error } = await supabase.from("usuarios").update({ nome: u.nome, email: u.email, perfil: u.perfil }).eq("id", u.id);
    if (error) { alert("Erro ao editar: " + error.message); return; }
    await carregarUsuarios();
  };

  const handleExcluirUsuario = async (id) => {
    const { error } = await supabase.from("usuarios").delete().eq("id", id);
    if (error) { alert("Erro ao excluir: " + error.message); return; }
    await carregarUsuarios();
  };

  const handleLogout = async () => {
    await supabase.auth.signOut();
    setUsuarioAtual(null);
  };

  const handleImportar = (csvContent, fileName) => {
    try {
      const lines = csvContent.trim().split("\n");
      const novasNotas = [];
      for (let i = 1; i < lines.length; i++) {
        const cols = lines[i].split(";");
        if (cols.length < 8) continue;
        const chave = cols[6]?.replace(/'/g, "").trim();
        if (!chave) continue;
        if (notas.find(n => n.chave === chave)) continue;

        const cnpjArquivo = cols[0]?.trim();
        const empMatch = empresas.find(e =>
          e.ativa && (
            cnpjArquivo === e.cnpj ||
            chave.startsWith((cnpjArquivo || "").replace(/[^0-9]/g,"").slice(0,14))
          )
        );

        novasNotas.push({
          id: `imp_${Date.now()}_${i}`,
          empresa: empMatch?.id || "",
          _cnpjDetectado: cnpjArquivo,
          _empresaId: empMatch?.id || "",
          fornecedor: cnpjArquivo,
          razaoSocial: cols[1]?.trim(),
          numNota: cols[2]?.trim(),
          cfop: cols[3]?.trim(),
          dtEmissao: cols[4]?.trim(),
          dtApresentacao: cols[5]?.trim() || "-",
          chave,
          valor: parseValor(cols[7]),
          qtdeDias: parseInt(cols[8]) || 0,
          status: "Identificada",
          centroCusto: "", codigoCusto: "", lancamentos: [], finalidade: "", responsavel: "", noRM: null, pagoReanalise: false, pagoDesembaraco: false, pagoIcms: false, pagoMulta10: false, pagoMulta: false, pagoJuros: false, dtSelada: "",
          taxaReanalise: 0, taxaDesembaraco: 0, icmsAntecipado: 0, multa10pct: 0, multa: 0, juros: 0,
          dtReanalise: "", dtDesembaraco: "", dtPostergacao: "", dtVencReanalise: "", dtVencDesembaraco: "", dtVencIcms: "", obs: "",
          dtImportacao: new Date().toLocaleDateString("pt-BR"),
          historico: []
        });
      }
      if (novasNotas.length === 0) {
        alert("Todas as notas do arquivo já estão no sistema.");
        return;
      }
      setImportPendente({ notas: novasNotas, fileName });
    } catch (e) {
      alert("Erro ao processar o arquivo. Verifique se é um CSV exportado do DTE.");
    }
  };

  const handleConfirmarImport = async (notasFinais) => {
    try {
      const registros = notasFinais.map(notaParaDb);
      const { error } = await supabase.from("notas").upsert(registros, { onConflict: "chave" });
      if (error) { alert("Erro ao importar: " + error.message); return; }
      const agora = new Date().toLocaleString("pt-BR");
      await supabase.from("configuracoes").upsert({ id: "global", ultima_importacao: agora, atualizado_em: new Date().toISOString() });
      setUltimaImportacao(agora);
      setImportPendente(null);
      await carregarNotas();
      alert(`✅ ${notasFinais.length} nota(s) importada(s) com sucesso!`);
    } catch (e) {
      alert("Erro inesperado: " + e.message);
    }
  };

  const notasAtivas = notas.filter(n => !["Desembaraçada", "Recusada"].includes(n.status));
  const criticas = notasAtivas.filter(n => n.qtdeDias >= 60).length;

  const navItems = [
    { id: "dashboard", label: "Acompanhamento", icon: "🏠" },
    { id: "gerencial", label: "Dashboard Gerencial", icon: "📊" },
    { id: "notas", label: "Painel de Notas", icon: "📋" },
    { id: "relatorios", label: "Relatórios", icon: "📈" },
    { id: "configuracoes", label: "Configurações", icon: "⚙️" },
  ];

  if (loginCarregando) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: "#1a4a4a" }}>
        <div className="text-center">
          <div className="w-16 h-16 rounded-xl flex items-center justify-center font-black text-white text-2xl mx-auto mb-4" style={{ background: "#E8450A" }}>E</div>
          <p className="text-sm mt-2" style={{ color: "#4db8b8" }}>Carregando...</p>
        </div>
      </div>
    );
  }

  if (!usuarioAtual) {
    return <TelaLogin onLogin={(user) => { setUsuarioAtual(user); }} logoUrl={logoUrl} />;
  }

  if (carregando) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: "#0d3535", fontFamily: "'Segoe UI', system-ui, sans-serif" }}>
        <div className="text-center">
          <div className="w-16 h-16 rounded-xl flex items-center justify-center font-black text-white text-2xl mx-auto mb-4" style={{ background: "#E8450A" }}>E</div>
          <p className="font-bold text-white text-lg">Enerwatt</p>
          <p className="text-sm mt-1" style={{ color: "#4db8b8" }}>Carregando sistema...</p>
          <div className="mt-4 w-48 h-1 rounded-full mx-auto overflow-hidden" style={{ background: "rgba(77,184,184,0.2)" }}>
            <div className="h-full rounded-full animate-pulse" style={{ background: "#E8450A", width: "60%" }}></div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen flex" style={{ fontFamily: "'Segoe UI', system-ui, sans-serif", background: "#edf5f5" }}>

      {/* SIDEBAR — petróleo escuro estilo imagem 3 */}
      <aside className="w-56 flex-shrink-0 hidden md:flex flex-col" style={{ background: "#1a4a4a", minHeight: "100vh", position: "relative", overflow: "hidden" }}>
        {/* Torres decorativas sutis */}
        <div className="absolute inset-0 pointer-events-none" style={{ overflow: "hidden" }}>
          <TorresSVG opacity={0.06} corTorre="#7ecece" />
        </div>
        {/* Logo topo */}
        <div className="relative z-10 px-5 py-4" style={{ borderBottom: "1px solid rgba(126,206,206,0.15)" }}>
          {logoUrl
            ? <img src={logoUrl} alt="Logo" className="h-9 object-contain" style={{ filter: "brightness(0) invert(1)" }} />
            : (
              <div>
                <p className="font-black text-lg text-white" style={{ fontFamily: "Georgia, serif", fontStyle: "italic" }}>Enerwatt</p>
                <p className="text-xs tracking-widest" style={{ color: "#E8450A" }}>ENGENHARIA</p>
              </div>
            )}
        </div>

        {/* Nav */}
        <nav className="relative z-10 flex-1 px-3 py-4 space-y-0.5">
          {navItems.map(item => (
            <button key={item.id} onClick={() => setTela(item.id)}
              className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-all text-left"
              style={tela === item.id
                ? { background: "#E8450A", color: "white" }
                : { color: "rgba(255,255,255,0.6)", background: "transparent" }}>
              <span className="text-base">{item.icon}</span>
              <span>{item.label}</span>
              {item.id === "dashboard" && criticas > 0 && (
                <span className="ml-auto text-xs font-bold px-1.5 py-0.5 rounded-full" style={{ background: tela === "dashboard" ? "rgba(255,255,255,0.3)" : "#c0392b", color: "white" }}>{criticas}</span>
              )}
            </button>
          ))}
        </nav>

        {/* Usuário */}
        <div className="relative z-10 px-3 py-4" style={{ borderTop: "1px solid rgba(126,206,206,0.15)" }}>
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-full flex items-center justify-center text-white text-xs font-bold flex-shrink-0" style={{ background: "#E8450A" }}>{usuarioAtual.nome.charAt(0)}</div>
            <div className="flex-1 min-w-0">
              <p className="text-xs font-semibold truncate text-white">{usuarioAtual.nome}</p>
              <p className="text-xs capitalize" style={{ color: "#7ecece" }}>{usuarioAtual.perfil}</p>
            </div>
            <button onClick={handleLogout} title="Sair" className="text-xs px-2 py-1 rounded font-semibold" style={{ color: "#7ecece", background: "rgba(126,206,206,0.1)" }}>Sair</button>
          </div>
        </div>
      </aside>

      {/* MAIN */}
      <main className="flex-1 flex flex-col overflow-hidden">
        {/* Header — petróleo com linha laranja embaixo, igual imagem 3 */}
        <header className="flex items-center justify-between px-6 py-0" style={{ background: "#1a4a4a", borderBottom: "3px solid #E8450A", minHeight: 52 }}>
          <div className="flex items-center gap-3">
            <div style={{ width: 3, height: 22, background: "#E8450A", borderRadius: 2 }} />
            <h1 className="font-bold text-sm text-white">
              {navItems.find(n => n.id === tela)?.label}
            </h1>
            <span className="text-xs hidden sm:block" style={{ color: "rgba(255,255,255,0.4)" }}>— Desembaraço Extemporâneo DTE/AM</span>
          </div>
          <div className="flex items-center gap-3">
            <span className="text-xs hidden sm:block" style={{ color: "rgba(255,255,255,0.5)" }}>{new Date().toLocaleDateString("pt-BR")}</span>
            {criticas > 0 && (
              <span className="text-xs font-bold px-3 py-1 rounded-full text-white" style={{ background: "#c0392b" }}>🔴 {criticas} crítica(s)</span>
            )}
          </div>
        </header>

        <div className="flex-1 overflow-y-auto p-6" style={{ background: "#edf5f5" }}>
          {tela === "dashboard" && <Dashboard notas={notas} onVerNota={n => setNotaSelecionada(n)} onIrParaPainel={() => setTela("notas")} />}
          {tela === "notas" && <PainelNotas notas={notas} onVerNota={n => setNotaSelecionada(n)} onImportar={handleImportar} ultimaImportacao={ultimaImportacao} empresas={empresas} />}
          {tela === "gerencial" && <DashboardGerencial notas={notas} empresas={empresas} />}
          {tela === "relatorios" && <Relatorios notas={notas} />}
          {tela === "configuracoes" && <Configuracoes usuarios={usuarios} onSalvarUsuario={handleSalvarUsuario} onEditarUsuario={handleEditarUsuario} onExcluirUsuario={handleExcluirUsuario} logoUrl={logoUrl} onSalvarLogo={handleSalvarLogo} empresas={empresas} onSalvarEmpresa={handleSalvarEmpresa} perfilAtual={usuarioAtual.perfil} />}
        </div>
      </main>

      {notaSelecionada && (
        <ModalNota nota={notaSelecionada} onClose={() => setNotaSelecionada(null)} onSave={handleSalvarNota} usuarioAtual={usuarioAtual} />
      )}

      {importPendente && (
        <ModalConfirmImport
          notasParaImportar={importPendente.notas}
          empresas={empresas}
          fileName={importPendente.fileName}
          usuarioNome={usuarioAtual.nome}
          onConfirmar={handleConfirmarImport}
          onCancelar={() => setImportPendente(null)}
        />
      )}
    </div>
  );
}
