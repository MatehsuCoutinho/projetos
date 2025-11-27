import React, { useState } from "react";
import "./App.css";
import { saveAs } from "file-saver";
import * as XLSX from "xlsx";

// =============================
// COMPONENTE PRINCIPAL
// =============================
export default function TravelSheetApp() {
  const [rows, setRows] = useState([]);
  const [baseCurrency, setBaseCurrency] = useState("BRL");
  const [exchangeRates, setExchangeRates] = useState({});

  function newRow() {
    return {
      name: "",
      origin: "",
      destination: "",
      date: "",
      airline: "",
      price: "",
      currency: baseCurrency,
      hotel: "",
      connections: [],
    };
  }

  function addRow() {
    setRows((r) => [...r, newRow()]);
  }

  function add22Rows() {
    const arr = Array.from({ length: 22 }, newRow);
    setRows(arr);
  }

  function updateRow(i, field, value) {
    setRows((rs) => {
      const copy = [...rs];
      copy[i] = { ...copy[i], [field]: value };
      return copy;
    });
  }

  function removeRow(i) {
    setRows((rs) => rs.filter((_, idx) => idx !== i));
  }

  function addConnection(i) {
    setRows((rs) => {
      const copy = [...rs];
      copy[i].connections.push({ origin: "", destination: "", airline: "", price: "" });
      return copy;
    });
  }

  function updateConnection(i, ci, field, value) {
    setRows((rs) => {
      const copy = [...rs];
      const conns = [...copy[i].connections];
      conns[ci] = { ...conns[ci], [field]: value };
      copy[i].connections = conns;
      return copy;
    });
  }

  function removeConnection(i, ci) {
    setRows((rs) => {
      const copy = [...rs];
      copy[i].connections = copy[i].connections.filter((_, idx) => idx !== ci);
      return copy;
    });
  }

  function setRate(currency, rate) {
    setExchangeRates((s) => ({ ...s, [currency]: parseFloat(rate) }));
  }

  // =============================
  // CONVERSÃO DE PREÇOS
  // =============================
  function parsePrice(p) {
    if (!p) return 0;
    const norm = String(p).replace(/\./g, "").replace(/,/, ".");
    const n = parseFloat(norm);
    return isNaN(n) ? 0 : n;
  }

  function totalPerRow(row) {
    const base = parsePrice(row.price);
    const hotel = parsePrice(row.hotel);

    const rowCurrency = row.currency || baseCurrency;
    const rate = exchangeRates[rowCurrency];
    const multiplier = rowCurrency === baseCurrency ? 1 : rate;
    if (!multiplier) return NaN;

    let sum = (base + hotel) * multiplier;

    // Somar conexões
    for (const c of row.connections) {
      const cp = parsePrice(c.price);
      sum += cp * multiplier;
    }

    return sum;
  }

  function grandTotal() {
    let total = 0;
    for (const r of rows) {
      const v = totalPerRow(r);
      if (isNaN(v)) return NaN;
      total += v;
    }
    return total;
  }

  // =============================
  // EXPORTAÇÃO EXCEL
  // =============================
function exportExcel() {
  const wsData = [];

  wsData.push([
  "Passageiro",       
  "Cidade Origem",   
  "Cidade Destino",   
  "Data Viagem",      
  "Companhia Aérea",  
  "Valor Voo (BRL)",  
  "Valor Hotel (3 noites)", 
  "Moeda",           
  "Total (BRL)",      
  "Conexões"         
  ]);

  for (const r of rows) {
    const converted = totalPerRow(r);
    const conexoesText = r.connections
      .map((c) => `${c.origin}→${c.destination} (${c.airline}) R$${c.price}`)
      .join(" | ");

    wsData.push([
      r.name,
      r.origin,
      r.destination,
      r.date,
      r.airline,
      r.price,
      r.hotel,
      r.currency,
      isNaN(converted) ? "CONVERTER" : converted,
      conexoesText,
    ]);
  }

  // =============================
  // LINHA DE TOTAL GERAL NO EXCEL
  // =============================
  const totalGeral = grandTotal();
  wsData.push([]);
  wsData.push([
    "", "", "", "", "", "", "",
    "TOTAL",
    isNaN(totalGeral) ? "CONVERTER" : totalGeral.toFixed(2),
    ""
  ]);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(wb, ws, "Viagem");

  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  saveAs(new Blob([buf]), "planilha_viagem.xlsx");
}

  return (
    <div className="app-container">
      <h1>Gerador de Planilha de Viagem</h1>

      <div className="toolbar">
        <button className="btn btn-blue" onClick={addRow}>Adicionar linha</button>
        <button className="btn btn-green" onClick={add22Rows}>Preencher 22 linhas</button>
        <button className="btn btn-indigo" onClick={exportExcel}>Exportar Excel</button>
      </div>

      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Nome</th>
            <th>Origem</th>
            <th>Destino</th>
            <th>Data</th>
            <th>Companhia</th>
            <th>Preço Voo</th>
            <th>Hotel (3 noites)</th>
            <th>Moeda</th>
            <th>Total ({baseCurrency})</th>
            <th>Conexões</th>
            <th>Ações</th>
          </tr>
        </thead>

        <tbody>
          {rows.map((row, i) => (
            <tr key={i}>
              <td>{i + 1}</td>
              <td><input value={row.name} onChange={(e) => updateRow(i, "name", e.target.value)} /></td>
              <td><input value={row.origin} onChange={(e) => updateRow(i, "origin", e.target.value)} /></td>
              <td><input value={row.destination} onChange={(e) => updateRow(i, "destination", e.target.value)} /></td>
              <td><input type="date" value={row.date} onChange={(e) => updateRow(i, "date", e.target.value)} /></td>
              <td><input value={row.airline} onChange={(e) => updateRow(i, "airline", e.target.value)} /></td>
              <td><input value={row.price} onChange={(e) => updateRow(i, "price", e.target.value)} /></td>
              <td><input value={row.hotel} onChange={(e) => updateRow(i, "hotel", e.target.value)} /></td>
              <td><input value={row.currency} onChange={(e) => updateRow(i, "currency", e.target.value.toUpperCase())} /></td>

              <td>{isNaN(totalPerRow(row)) ? "converter" : totalPerRow(row).toFixed(2)}</td>

              <td>
                {row.connections.map((c, ci) => (
                  <div key={ci} className="connection-box">
                    <input placeholder="Origem" value={c.origin} onChange={(e) => updateConnection(i, ci, "origin", e.target.value)} />
                    <input placeholder="Destino" value={c.destination} onChange={(e) => updateConnection(i, ci, "destination", e.target.value)} />
                    <input placeholder="Companhia" value={c.airline} onChange={(e) => updateConnection(i, ci, "airline", e.target.value)} />
                    <input placeholder="Preço" value={c.price} onChange={(e) => updateConnection(i, ci, "price", e.target.value)} />
                    <button className="btn btn-red" onClick={() => removeConnection(i, ci)}>X</button>
                  </div>
                ))}

                <button className="btn btn-blue" onClick={() => addConnection(i)}>+ Conexão</button>
              </td>

              <td>
                <button className="btn btn-red" onClick={() => removeRow(i)}>Remover</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      <div className="total-box">
        <strong>Total ({baseCurrency}): </strong>
        {isNaN(grandTotal()) ? "Há moedas sem taxa informada" : grandTotal().toFixed(2)}
      </div>
    </div>
  );
}