import { useMemo, useState } from 'react';
import { Platform, ScrollView, Text, TextInput, TouchableOpacity, View } from 'react-native';
import * as FileSystem from 'expo-file-system';
import * as Print from 'expo-print';
import * as Sharing from 'expo-sharing';
import * as XLSX from 'xlsx';
import { C } from './ui';
import { calc, dmy, inr, today } from './utils';

const MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

/* ---- Section 1: Customer Ledger Report Format ---- */
export function LedgerReportFormat({ db }) {
  const [searchQ, setSearchQ] = useState('');
  const [filterYear, setFilterYear] = useState('');
  const [filterMonth, setFilterMonth] = useState('');
  const loans = db.loans || [];
  const customers = db.customers || [];
  const payments = db.payments || [];

  const years = useMemo(() => {
    const ys = new Set(loans.map((l) => (l.date || '').slice(0, 4)).filter(Boolean));
    return Array.from(ys).sort().reverse();
  }, [loans]);

  const rows = useMemo(() => loans.filter((l) => {
    if (filterYear && !(l.date || '').startsWith(filterYear)) return false;
    if (filterMonth && (l.date || '').slice(5, 7) !== filterMonth) return false;
    const cust = customers.find((c) => c.id === l.customerId) || {};
    const q = searchQ.toLowerCase();
    if (q && !(cust.name || '').toLowerCase().includes(q) && !(cust.phone || '').includes(q)) return false;
    return true;
  }).map((l) => {
    const cust = customers.find((c) => c.id === l.customerId) || {};
    const lPays = payments.filter((p) => p.loanId === l.id);
    const c = calc(l, payments);
    const loanMonth = (l.date || '').slice(0, 7);
    const prevPays = lPays.filter((p) => (p.date || '') < loanMonth + '-01');
    const prevBalance = prevPays.reduce((s, p) => s + p.amount, 0);
    return {
      date: l.date || '', loanNo: l.no,
      nameKannada: cust.nameKannada || cust.name || '-',
      interest: Math.round(c.interest),
      prevMonthBalance: Math.round(prevBalance),
      mobile: cust.phone || '-',
      oldBalance: Math.round(Math.max(0, l.principal - prevBalance)),
    };
  }), [loans, customers, payments, searchQ, filterYear, filterMonth]);

  const exportExcel = () => {
    const shop = (db.settings || {}).shopName || 'Kuber Jewellery';
    const wb = XLSX.utils.book_new();
    const hdr = ['Date','Name (Kannada)','Interest','Previous Month Balance','Mobile Number','Old Balance'];
    const data = rows.map((r) => [dmy(r.date), r.nameKannada, r.interest, r.prevMonthBalance, r.mobile, r.oldBalance]);
    const ws = XLSX.utils.aoa_to_sheet([[shop + ' - Customer Ledger Report'],['Generated: ' + dmy(today())],[],hdr,...data]);
    ws['!cols'] = [{wch:14},{wch:24},{wch:12},{wch:22},{wch:16},{wch:14}];
    for (let ci = 0; ci < 6; ci++) {
      const ref = XLSX.utils.encode_cell({ r: 3, c: ci });
      if (ws[ref]) ws[ref].s = { font: { bold: true }, alignment: { horizontal: 'center' } };
    }
    for (let ri = 3; ri < 4 + data.length; ri++) {
      for (let ci = 0; ci < 6; ci++) {
        const ref = XLSX.utils.encode_cell({ r: ri, c: ci });
        if (!ws[ref]) ws[ref] = { v: '', t: 's' };
        ws[ref].s = { ...(ws[ref].s || {}), border: { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} } };
      }
    }
    XLSX.utils.book_append_sheet(wb, ws, 'Ledger Report');
    if (Platform.OS === 'web') { XLSX.writeFile(wb, 'customer_ledger_report.xlsx'); }
    else {
      const wbout = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
      const path = FileSystem.cacheDirectory + 'customer_ledger_report.xlsx';
      FileSystem.writeAsStringAsync(path, wbout, { encoding: FileSystem.EncodingType.Base64 })
        .then(() => Sharing.shareAsync(path, { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', dialogTitle: 'Export Excel' }));
    }
  };

  const exportPDF = async () => {
    const shop = (db.settings || {}).shopName || 'Kuber Jewellery';
    const tr = rows.map((r, i) =>
      `<tr style="background:${i%2===0?'#fff':'#F7FAFF'}"><td>${dmy(r.date)}</td><td>${r.nameKannada}</td><td align="right">${inr(r.interest)}</td><td align="right">${inr(r.prevMonthBalance)}</td><td>${r.mobile}</td><td align="right">${inr(r.oldBalance)}</td></tr>`
    ).join('');
    const html = `<html><head><meta charset="utf-8"><style>body{font-family:sans-serif;padding:16px;font-size:12px}h2{color:#12263A}table{width:100%;border-collapse:collapse}th{background:#12263A;color:#fff;padding:7px 8px;text-align:left;font-size:11px}td{padding:6px 8px;border:1px solid #ddd}.tot{background:#D6E8FF;font-weight:bold}</style></head><body><h2>${shop} - Customer Ledger Report</h2><p>Generated: ${dmy(today())}</p><table><thead><tr><th>Date</th><th>Name (Kannada)</th><th>Interest</th><th>Prev Month Bal</th><th>Mobile</th><th>Old Balance</th></tr></thead><tbody>${tr}</tbody><tfoot><tr class="tot"><td colspan="2">Total (${rows.length})</td><td align="right">${inr(rows.reduce((s,r)=>s+r.interest,0))}</td><td align="right">${inr(rows.reduce((s,r)=>s+r.prevMonthBalance,0))}</td><td></td><td align="right">${inr(rows.reduce((s,r)=>s+r.oldBalance,0))}</td></tr></tfoot></table></body></html>`;
    if (Platform.OS === 'web') { const w = window.open('','_blank'); w.document.write(html); w.document.close(); w.print(); }
    else { const { uri } = await Print.printToFileAsync({ html }); await Sharing.shareAsync(uri, { mimeType: 'application/pdf', dialogTitle: 'Export PDF' }); }
  };

  return (
    <View style={{ flex: 1, backgroundColor: '#fff' }}>
      <View style={{ backgroundColor: '#F7FAFF', borderBottomWidth: 1, borderColor: '#D6E8FF', padding: 12 }}>
        <View style={{ flexDirection: 'row', alignItems: 'center', backgroundColor: '#fff', borderRadius: 10, borderWidth: 1, borderColor: C.line, paddingHorizontal: 10, marginBottom: 8 }}>
          <Text style={{ color: C.mute, fontSize: 15, marginRight: 6 }}>🔍</Text>
          <TextInput value={searchQ} onChangeText={setSearchQ} placeholder="Search by name or mobile"
            placeholderTextColor="#9AA7B4" style={{ flex: 1, paddingVertical: 9, fontSize: 14, color: C.text }} />
        </View>
        <ScrollView horizontal showsHorizontalScrollIndicator={false} style={{ marginBottom: 6 }}>
          <View style={{ flexDirection: 'row', gap: 6 }}>
            {['', ...years].map((y) => (
              <TouchableOpacity key={y || 'all'} onPress={() => { setFilterYear(y); setFilterMonth(''); }}
                style={{ paddingHorizontal: 12, paddingVertical: 6, borderRadius: 18, borderWidth: 1,
                  borderColor: filterYear === y ? C.navy : C.line, backgroundColor: filterYear === y ? C.navy : '#fff' }}>
                <Text style={{ color: filterYear === y ? '#fff' : C.text, fontWeight: '600', fontSize: 12 }}>{y || 'All Years'}</Text>
              </TouchableOpacity>
            ))}
          </View>
        </ScrollView>
        {filterYear ? (
          <ScrollView horizontal showsHorizontalScrollIndicator={false}>
            <View style={{ flexDirection: 'row', gap: 6 }}>
              {['', ...MN.map((_, i) => String(i + 1).padStart(2, '0'))].map((m, idx) => (
                <TouchableOpacity key={m || 'all'} onPress={() => setFilterMonth(m)}
                  style={{ paddingHorizontal: 10, paddingVertical: 5, borderRadius: 18, borderWidth: 1,
                    borderColor: filterMonth === m ? C.gold : C.line, backgroundColor: filterMonth === m ? C.gold : '#fff' }}>
                  <Text style={{ color: filterMonth === m ? '#fff' : C.text, fontSize: 12 }}>{m ? MN[idx - 1] : 'All Months'}</Text>
                </TouchableOpacity>
              ))}
            </View>
          </ScrollView>
        ) : null}
      </View>
      <View style={{ flexDirection: 'row', alignItems: 'center', paddingHorizontal: 12, paddingVertical: 8, backgroundColor: '#F7FAFF', borderBottomWidth: 1, borderColor: '#D6E8FF', gap: 6 }}>
        <Text style={{ flex: 1, color: C.mute, fontSize: 12 }}>{rows.length} records</Text>
        <TouchableOpacity onPress={exportExcel} style={{ backgroundColor: C.green, borderRadius: 8, paddingHorizontal: 10, paddingVertical: 6 }}>
          <Text style={{ color: '#fff', fontWeight: '700', fontSize: 12 }}>📥 Excel</Text>
        </TouchableOpacity>
        <TouchableOpacity onPress={exportPDF} style={{ backgroundColor: C.red, borderRadius: 8, paddingHorizontal: 10, paddingVertical: 6 }}>
          <Text style={{ color: '#fff', fontWeight: '700', fontSize: 12 }}>🖨 PDF</Text>
        </TouchableOpacity>
      </View>
      <ScrollView style={{ flex: 1 }} contentContainerStyle={{ flexGrow: 1 }}>
        <ScrollView horizontal showsHorizontalScrollIndicator={true}>
          <View>
            <View style={{ flexDirection: 'row', backgroundColor: C.navy, paddingVertical: 10 }}>
              {[['Date',90],['Name (Kannada)',160],['Interest',90],['Prev Month Bal',120],['Mobile',120],['Old Balance',100]].map(([h,w]) => (
                <Text key={h} style={{ width: w, color: '#fff', fontWeight: '700', fontSize: 11, paddingHorizontal: 6, textAlign: 'center' }}>{h}</Text>
              ))}
            </View>
            {rows.length === 0
              ? <View style={{ paddingVertical: 40, alignItems: 'center', width: 680 }}><Text style={{ color: C.mute }}>No records found.</Text></View>
              : rows.map((r, i) => (
                <View key={r.loanNo + i} style={{ flexDirection: 'row', borderBottomWidth: 1, borderColor: '#E8EDF2', backgroundColor: i % 2 === 0 ? '#fff' : '#F7FAFF', paddingVertical: 9 }}>
                  <Text style={{ width: 90, fontSize: 12, color: C.text, paddingHorizontal: 6, textAlign: 'center' }}>{dmy(r.date)}</Text>
                  <Text style={{ width: 160, fontSize: 12, color: C.text, paddingHorizontal: 6, fontWeight: '600' }} numberOfLines={1}>{r.nameKannada}</Text>
                  <Text style={{ width: 90, fontSize: 12, color: C.gold, paddingHorizontal: 6, textAlign: 'right', fontWeight: '600' }}>{inr(r.interest)}</Text>
                  <Text style={{ width: 120, fontSize: 12, color: C.navy, paddingHorizontal: 6, textAlign: 'right' }}>{inr(r.prevMonthBalance)}</Text>
                  <Text style={{ width: 120, fontSize: 12, color: C.text, paddingHorizontal: 6 }}>{r.mobile}</Text>
                  <Text style={{ width: 100, fontSize: 12, color: C.red, paddingHorizontal: 6, textAlign: 'right', fontWeight: '600' }}>{inr(r.oldBalance)}</Text>
                </View>
              ))
            }
            {rows.length > 0 && (
              <View style={{ flexDirection: 'row', backgroundColor: '#D6E8FF', paddingVertical: 10 }}>
                <Text style={{ width: 90, fontSize: 12, fontWeight: '800', color: C.navy, paddingHorizontal: 6 }}>TOTAL</Text>
                <Text style={{ width: 160, paddingHorizontal: 6 }} />
                <Text style={{ width: 90, fontSize: 12, fontWeight: '800', color: C.gold, paddingHorizontal: 6, textAlign: 'right' }}>{inr(rows.reduce((s,r)=>s+r.interest,0))}</Text>
                <Text style={{ width: 120, fontSize: 12, fontWeight: '800', color: C.navy, paddingHorizontal: 6, textAlign: 'right' }}>{inr(rows.reduce((s,r)=>s+r.prevMonthBalance,0))}</Text>
                <Text style={{ width: 120, paddingHorizontal: 6 }} />
                <Text style={{ width: 100, fontSize: 12, fontWeight: '800', color: C.red, paddingHorizontal: 6, textAlign: 'right' }}>{inr(rows.reduce((s,r)=>s+r.oldBalance,0))}</Text>
              </View>
            )}
          </View>
        </ScrollView>
      </ScrollView>
    </View>
  );
}

/* ---- Loan detail table for YearwiseKannadaReport ---- */
function LoanDetailTable({ l, name }) {
  const rows = [
    ['ಹೆಸರು (Name)',        name || '-',          '#12263A'],
    ['ದಿನಾಂಕ (Date)',       dmy(l.loanDate),      '#12263A'],
    ['ಸಂಖ್ಯೆ (Number)',     l.loanNo,             '#12263A'],
    ['ಮೊತ್ತ (Amount)',      inr(l.principal),     '#12263A'],
    ['ಬಡ್ಡಿ (Interest)',    inr(l.interest),      '#B8912F'],
    ['ತಿಂಗಳ ಕಂತು (Due)',   l.dueDate ? dmy(l.dueDate) : '-', '#6B7A8A'],
    ['ಹಳೆ ಬಾಕಿ (Old Bal)', inr(l.oldBalance),    '#B93A2E'],
    ['ಹಿಂದಿನ ತಿಂಗಳ ಬಾಕಿ',  inr(l.yesterdayMonthBal), '#1E8A58'],
  ];
  return (
    <View style={{ marginTop: 6, borderRadius: 8, overflow: 'hidden', borderWidth: 1, borderColor: '#E8EDF2' }}>
      {rows.map(([label, value, col], i) => (
        <View key={label} style={{ flexDirection: 'row', backgroundColor: i % 2 === 0 ? '#F7FAFF' : '#fff', paddingVertical: 7, paddingHorizontal: 10 }}>
          <Text style={{ flex: 1, fontSize: 11, color: '#6B7A8A' }}>{label}</Text>
          <Text style={{ fontSize: 12, fontWeight: '700', color: col }}>{value}</Text>
        </View>
      ))}
    </View>
  );
}

/* ---- Section 2: Year-wise Member Report (Table View) ---- */
export function YearwiseKannadaReport({ db }) {
  const now = new Date();
  const [selYear, setSelYear] = useState(String(now.getFullYear()));
  const [selMonth, setSelMonth] = useState('');
  const [selDate, setSelDate] = useState('');
  const [searchQ, setSearchQ] = useState('');
  const loans = db.loans || [];
  const customers = db.customers || [];
  const payments = db.payments || [];

  const years = useMemo(() => {
    const ys = new Set(loans.map((l) => (l.date || '').slice(0, 4)).filter(Boolean));
    ys.add(String(now.getFullYear()));
    return Array.from(ys).sort().reverse();
  }, [loans]);

  const flatRows = useMemo(() => {
    const q = searchQ.toLowerCase();
    const result = [];
    loans.forEach((l) => {
      if (!(l.date || '').startsWith(selYear)) return;
      if (selMonth && (l.date || '').slice(5, 7) !== selMonth) return;
      if (selDate && (l.date || '') !== selDate) return;
      const c = customers.find((x) => x.id === l.customerId) || {};
      if (q && !(c.name || '').toLowerCase().includes(q) &&
          !(c.nameKannada || '').toLowerCase().includes(q) &&
          !(c.phone || '').includes(q) &&
          !(c.custNo || '').toLowerCase().includes(q)) return;
      const cv = calc(l, payments);
      const lPays = payments.filter((p) => p.loanId === l.id);
      const paid = lPays.reduce((s, p) => s + p.amount, 0);
      const prevPays = lPays.filter((p) => (p.date || '') < (l.date || ''));
      const prevBal = prevPays.reduce((s, p) => s + p.amount, 0);
      const monthlyDue = Math.round((l.principal * (l.rate || 0)) / 100);
      result.push({
        date: l.date || '',
        custName: c.nameKannada || c.name || '-',
        principal: l.principal || 0,
        interest: Math.round(cv.interest),
        monthlyDue,
        mobile: c.phone || '-',
        custNo: c.custNo || '-',
        prevBalance: Math.round(prevBal),
        amountPaid: Math.round(paid),
        pending: Math.round(cv.outstanding),
        totPaid: Math.round(paid),
        totInterest: Math.round(cv.interest),
        status: l.status || 'active',
        remarks: l.remarks || '',
        loanNo: l.no,
      });
    });
    result.sort((a, b) => a.date.localeCompare(b.date));
    return result;
  }, [loans, customers, payments, selYear, selMonth, selDate, searchQ]);

  const grandPrincipal = flatRows.reduce((s, r) => s + r.principal, 0);
  const grandInterest = flatRows.reduce((s, r) => s + r.interest, 0);
  const grandPaid = flatRows.reduce((s, r) => s + r.amountPaid, 0);
  const grandPending = flatRows.reduce((s, r) => s + r.pending, 0);

  const periodLabel = selMonth ? MN[parseInt(selMonth, 10) - 1] + ' ' + selYear : selYear;

  const exportPDF = async () => {
    const shop = (db.settings || {}).shopName || 'Kuber Jewellery';
    const tr = flatRows.map((r, i) => `<tr style="background:${i%2===0?'#fff':'#F7FAFF'}">
      <td>${dmy(r.date)}</td><td>${r.custName}</td>
      <td align="right">${inr(r.principal)}</td><td align="right">${inr(r.interest)}</td>
      <td align="right">${inr(r.monthlyDue)}</td><td>${r.mobile}</td>
      <td>${r.custNo}</td><td align="right">${inr(r.prevBalance)}</td>
      <td align="right">${inr(r.amountPaid)}</td><td align="right">${inr(r.pending)}</td>
      <td align="right">${inr(r.totPaid)}</td><td align="right">${inr(r.totInterest)}</td>
      <td>${r.status === 'active' ? 'Active' : 'Closed'}</td><td>${r.remarks}</td></tr>`).join('');
    const html = `<html><head><meta charset="utf-8"><style>body{font-family:sans-serif;padding:12px;font-size:10px}h2{color:#12263A;margin:0}h3{color:#B8912F;margin:4px 0 10px}table{width:100%;border-collapse:collapse}th{background:#12263A;color:#fff;padding:5px 6px;font-size:9px;white-space:nowrap}td{padding:4px 6px;border:1px solid #ddd;white-space:nowrap}.tot{background:#D6E8FF;font-weight:bold}</style></head><body><h2>${shop}</h2><h3>ವರ್ಷವಾರು ಸದಸ್ಯರ ವರದಿ &mdash; ${periodLabel}</h3><table><thead><tr><th>Date</th><th>Customer Name</th><th>Principal</th><th>Interest</th><th>Monthly Due</th><th>Mobile</th><th>Cust No</th><th>Prev Balance</th><th>Amt Paid</th><th>Pending</th><th>Total Paid</th><th>Total Interest</th><th>Status</th><th>Remarks</th></tr></thead><tbody>${tr}</tbody><tfoot><tr class="tot"><td colspan="2">TOTAL (${flatRows.length})</td><td align="right">${inr(grandPrincipal)}</td><td align="right">${inr(grandInterest)}</td><td></td><td></td><td></td><td></td><td align="right">${inr(grandPaid)}</td><td align="right">${inr(grandPending)}</td><td align="right">${inr(grandPaid)}</td><td align="right">${inr(grandInterest)}</td><td></td><td></td></tr></tfoot></table></body></html>`;
    if (Platform.OS === 'web') { const w = window.open('', '_blank'); w.document.write(html); w.document.close(); w.print(); }
    else { const { uri } = await Print.printToFileAsync({ html }); await Sharing.shareAsync(uri, { mimeType: 'application/pdf', dialogTitle: 'Export PDF' }); }
  };

  const exportExcel = () => {
    const shop = (db.settings || {}).shopName || 'Kuber Jewellery';
    const wb = XLSX.utils.book_new();
    const hdr = ['Date','Customer Name','Principal','Interest','Monthly Due','Mobile','Customer No','Prev Balance','Amount Paid','Pending','Total Paid','Total Interest','Status','Remarks'];
    const data = flatRows.map((r) => [dmy(r.date), r.custName, r.principal, r.interest, r.monthlyDue, r.mobile, r.custNo, r.prevBalance, r.amountPaid, r.pending, r.totPaid, r.totInterest, r.status, r.remarks]);
    const tag = selMonth ? MN[parseInt(selMonth,10)-1]+'_'+selYear : selYear;
    const ws = XLSX.utils.aoa_to_sheet([[shop+' - ವರ್ಷವಾರು ಸದಸ್ಯರ ವರದಿ - '+periodLabel],['ದಿನಾಂಕ: '+dmy(today())],[],hdr,...data,['ಒಟ್ಟು','',grandPrincipal,grandInterest,'','','','',grandPaid,grandPending,grandPaid,grandInterest,'','']]);
    ws['!cols'] = [{wch:12},{wch:22},{wch:12},{wch:12},{wch:12},{wch:14},{wch:12},{wch:14},{wch:12},{wch:12},{wch:12},{wch:14},{wch:10},{wch:20}];
    XLSX.utils.book_append_sheet(wb, ws, tag.slice(0,31));
    if (Platform.OS === 'web') { XLSX.writeFile(wb, 'varshavaru_sadasya_varadi_'+tag+'.xlsx'); }
    else {
      const wbout = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
      const path = FileSystem.cacheDirectory + 'varshavaru_sadasya_varadi.xlsx';
      FileSystem.writeAsStringAsync(path, wbout, { encoding: FileSystem.EncodingType.Base64 })
        .then(() => Sharing.shareAsync(path, { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', dialogTitle: 'Export Excel' }));
    }
  };

  const COLS = [
    { key: 'date',        label: 'Date',           w: 96,  align: 'center' },
    { key: 'custName',    label: 'Customer Name',  w: 160, align: 'left'   },
    { key: 'principal',   label: 'Principal',      w: 100, align: 'right'  },
    { key: 'interest',    label: 'Interest',       w: 100, align: 'right'  },
    { key: 'monthlyDue',  label: 'Monthly Due',    w: 100, align: 'right'  },
    { key: 'mobile',      label: 'Mobile',         w: 120, align: 'left'   },
    { key: 'custNo',      label: 'Cust No',        w: 100, align: 'center' },
    { key: 'prevBalance', label: 'Prev Balance',   w: 110, align: 'right'  },
    { key: 'amountPaid',  label: 'Amt Paid',       w: 100, align: 'right'  },
    { key: 'pending',     label: 'Pending',        w: 100, align: 'right'  },
    { key: 'totPaid',     label: 'Total Paid',     w: 100, align: 'right'  },
    { key: 'totInterest', label: 'Total Interest', w: 110, align: 'right'  },
    { key: 'status',      label: 'Status',         w: 80,  align: 'center' },
    { key: 'remarks',     label: 'Remarks',        w: 140, align: 'left'   },
  ];
  const totalW = COLS.reduce((s, c) => s + c.w, 0);

  const cellVal = (r, key) => {
    if (key === 'date') return dmy(r.date);
    if (key === 'status') return r.status === 'active' ? '● Active' : '✓ Closed';
    if (['principal','interest','monthlyDue','prevBalance','amountPaid','pending','totPaid','totInterest'].includes(key)) return inr(r[key]);
    return r[key] || '-';
  };
  const cellColor = (r, key) => {
    if (key === 'interest' || key === 'totInterest') return '#B8912F';
    if (key === 'pending') return '#B93A2E';
    if (key === 'amountPaid' || key === 'totPaid') return '#1E8A58';
    if (key === 'status') return r.status === 'active' ? '#1E8A58' : '#6B7A8A';
    return '#1B2733';
  };
  const totVal = (key) => {
    if (key === 'principal') return inr(grandPrincipal);
    if (key === 'interest' || key === 'totInterest') return inr(grandInterest);
    if (key === 'amountPaid' || key === 'totPaid') return inr(grandPaid);
    if (key === 'pending') return inr(grandPending);
    if (key === 'custName') return `${flatRows.length} records`;
    return '';
  };

  return (
    <View style={{ flex: 1, backgroundColor: '#F4F7FB' }}>
      {/* Header */}
      <View style={{ backgroundColor: '#12263A', paddingHorizontal: 16, paddingVertical: 12 }}>
        <Text style={{ color: '#B8912F', fontWeight: '800', fontSize: 16 }}>Kuber Jewellery – {selYear} Year-wise Member Report</Text>
        <Text style={{ color: '#B9C6D3', fontSize: 11, marginTop: 2 }}>ವರ್ಷವಾರು ಸದಸ್ಯರ ವರದಿ · {periodLabel}</Text>
      </View>
      {/* Filters */}
      <View style={{ backgroundColor: '#F7FAFF', borderBottomWidth: 1, borderColor: '#D6E8FF', padding: 10, gap: 8 }}>
        {/* Year pills */}
        <ScrollView horizontal showsHorizontalScrollIndicator={false}>
          <View style={{ flexDirection: 'row', gap: 6 }}>
            {years.map((y) => (
              <TouchableOpacity key={y} onPress={() => { setSelYear(y); setSelMonth(''); setSelDate(''); }}
                style={{ paddingHorizontal: 14, paddingVertical: 7, borderRadius: 18, borderWidth: 1,
                  borderColor: selYear === y ? '#B8912F' : '#D0DCF0', backgroundColor: selYear === y ? '#B8912F' : '#fff' }}>
                <Text style={{ color: selYear === y ? '#fff' : '#1B2733', fontWeight: '700', fontSize: 13 }}>{y}</Text>
              </TouchableOpacity>
            ))}
          </View>
        </ScrollView>
        {/* Month pills */}
        <ScrollView horizontal showsHorizontalScrollIndicator={false}>
          <View style={{ flexDirection: 'row', gap: 6 }}>
            {['', ...MN.map((_, i) => String(i + 1).padStart(2, '0'))].map((m, idx) => (
              <TouchableOpacity key={m || 'all'} onPress={() => { setSelMonth(m); setSelDate(''); }}
                style={{ paddingHorizontal: 10, paddingVertical: 5, borderRadius: 18, borderWidth: 1,
                  borderColor: selMonth === m ? C.navy : '#D0DCF0', backgroundColor: selMonth === m ? C.navy : '#fff' }}>
                <Text style={{ color: selMonth === m ? '#fff' : '#1B2733', fontSize: 12 }}>{m ? MN[idx - 1] : 'All Months'}</Text>
              </TouchableOpacity>
            ))}
          </View>
        </ScrollView>
        {/* Search + Date filter row */}
        <View style={{ flexDirection: 'row', gap: 8 }}>
          <View style={{ flex: 1, flexDirection: 'row', alignItems: 'center', backgroundColor: '#fff', borderRadius: 10, borderWidth: 1, borderColor: '#E1E6EC', paddingHorizontal: 10 }}>
            <Text style={{ color: '#6B7A8A', fontSize: 14, marginRight: 6 }}>🔍</Text>
            <TextInput value={searchQ} onChangeText={setSearchQ}
              placeholder="Name / Mobile / Cust No"
              placeholderTextColor="#9AA7B4"
              style={{ flex: 1, paddingVertical: 8, fontSize: 13, color: '#1B2733' }} />
            {searchQ ? <TouchableOpacity onPress={() => setSearchQ('')}><Text style={{ color: '#B93A2E', fontSize: 16, paddingLeft: 6 }}>✕</Text></TouchableOpacity> : null}
          </View>
          <View style={{ flexDirection: 'row', alignItems: 'center', backgroundColor: '#fff', borderRadius: 10, borderWidth: 1, borderColor: '#E1E6EC', paddingHorizontal: 10 }}>
            <TextInput value={selDate} onChangeText={setSelDate}
              placeholder="YYYY-MM-DD"
              placeholderTextColor="#9AA7B4"
              style={{ width: 110, paddingVertical: 8, fontSize: 12, color: '#1B2733' }} />
            {selDate ? <TouchableOpacity onPress={() => setSelDate('')}><Text style={{ color: '#B93A2E', fontSize: 16 }}>✕</Text></TouchableOpacity> : null}
          </View>
        </View>
        {/* Toolbar */}
        <View style={{ flexDirection: 'row', alignItems: 'center', gap: 6 }}>
          <Text style={{ flex: 1, color: '#6B7A8A', fontSize: 12 }}>{flatRows.length} records · {periodLabel}</Text>
          <TouchableOpacity onPress={exportExcel} style={{ backgroundColor: '#1E8A58', borderRadius: 8, paddingHorizontal: 12, paddingVertical: 7 }}>
            <Text style={{ color: '#fff', fontWeight: '700', fontSize: 12 }}>📥 Excel</Text>
          </TouchableOpacity>
          <TouchableOpacity onPress={exportPDF} style={{ backgroundColor: '#B93A2E', borderRadius: 8, paddingHorizontal: 12, paddingVertical: 7 }}>
            <Text style={{ color: '#fff', fontWeight: '700', fontSize: 12 }}>🖨 PDF</Text>
          </TouchableOpacity>
        </View>
      </View>

      {/* Table */}
      {flatRows.length === 0 ? (
        <View style={{ flex: 1, alignItems: 'center', justifyContent: 'center' }}>
          <Text style={{ color: '#6B7A8A', fontSize: 14 }}>No records for {periodLabel}.</Text>
        </View>
      ) : (
        <ScrollView style={{ flex: 1 }} contentContainerStyle={{ flexGrow: 1 }}>
          <ScrollView horizontal showsHorizontalScrollIndicator={true}>
            <View style={{ minWidth: totalW }}>
              {/* Column Headers */}
              <View style={{ flexDirection: 'row', backgroundColor: '#12263A', paddingVertical: 10 }}>
                {COLS.map((col) => (
                  <Text key={col.key} style={{ width: col.w, color: '#fff', fontWeight: '700', fontSize: 11,
                    paddingHorizontal: 6, textAlign: col.align }}>{col.label}</Text>
                ))}
              </View>
              {/* Data Rows */}
              {flatRows.map((r, i) => (
                <View key={r.loanNo + i} style={{ flexDirection: 'row', borderBottomWidth: 1,
                  borderColor: '#E8EDF2', backgroundColor: i % 2 === 0 ? '#fff' : '#F7FAFF', paddingVertical: 9 }}>
                  {COLS.map((col) => (
                    <Text key={col.key} numberOfLines={1}
                      style={{ width: col.w, fontSize: 12, paddingHorizontal: 6,
                        textAlign: col.align, color: cellColor(r, col.key),
                        fontWeight: ['custName','principal'].includes(col.key) ? '600' : '400' }}>
                      {cellVal(r, col.key)}
                    </Text>
                  ))}
                </View>
              ))}
              {/* Totals Row */}
              <View style={{ flexDirection: 'row', backgroundColor: '#D6E8FF', paddingVertical: 10 }}>
                {COLS.map((col) => (
                  <Text key={col.key} style={{ width: col.w, fontSize: 12, fontWeight: '800',
                    paddingHorizontal: 6, textAlign: col.align,
                    color: col.key === 'interest' || col.key === 'totInterest' ? '#B8912F'
                         : col.key === 'pending' ? '#B93A2E'
                         : col.key === 'amountPaid' || col.key === 'totPaid' ? '#1E8A58'
                         : '#12263A' }}>
                    {totVal(col.key)}
                  </Text>
                ))}
              </View>
            </View>
          </ScrollView>
        </ScrollView>
      )}
    </View>
  );
}
