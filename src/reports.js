import { useMemo, useState } from 'react';
import { Modal, Platform, ScrollView, Text, TextInput, TouchableOpacity, View } from 'react-native';
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
      <ScrollView horizontal style={{ flex: 1 }}>
        <View>
          <View style={{ flexDirection: 'row', backgroundColor: C.navy, paddingVertical: 10 }}>
            {[['Date',90],['Name (Kannada)',160],['Interest',90],['Prev Month Bal',120],['Mobile',120],['Old Balance',100]].map(([h,w]) => (
              <Text key={h} style={{ width: w, color: '#fff', fontWeight: '700', fontSize: 11, paddingHorizontal: 6, textAlign: 'center' }}>{h}</Text>
            ))}
          </View>
          <ScrollView>
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
          </ScrollView>
        </View>
      </ScrollView>
    </View>
  );
}

/* ---- Section 2: Year-wise Kannada Member Report ---- */
export function YearwiseKannadaReport({ db }) {
  const now = new Date();
  const [selYear, setSelYear] = useState(String(now.getFullYear()));
  const [searchQ, setSearchQ] = useState('');
  const [viewMode, setViewMode] = useState(false);
  const loans = db.loans || [];
  const customers = db.customers || [];
  const payments = db.payments || [];

  const years = useMemo(() => {
    const ys = new Set(loans.map((l) => (l.date || '').slice(0, 4)).filter(Boolean));
    const cur = String(now.getFullYear());
    ys.add(cur);
    return Array.from(ys).sort().reverse();
  }, [loans]);

  const custRows = useMemo(() => {
    return customers
      .filter((c) => {
        const q = searchQ.toLowerCase();
        if (q && !(c.name || '').toLowerCase().includes(q) && !(c.phone || '').includes(q)) return false;
        return loans.some((l) => l.customerId === c.id && (l.date || '').startsWith(selYear));
      })
      .map((c) => {
        const displayName = c.nameKannada || c.name;
        const cLoans = loans.filter((l) => l.customerId === c.id && (l.date || '').startsWith(selYear));
        const loanDetails = cLoans.map((l) => {
          const lPays = payments.filter((p) => p.loanId === l.id);
          const cv = calc(l, payments);
          const paid = lPays.reduce((s, p) => s + p.amount, 0);
          const prevPays = payments.filter((p) => p.loanId === l.id && (p.date || '') < (l.date || ''));
          const prevBal = prevPays.reduce((s, p) => s + p.amount, 0);
          return {
            loanNo: l.no,
            loanDate: l.date || '',
            principal: l.principal,
            interest: Math.round(cv.interest),
            prevBalance: Math.round(prevBal),
            paid: Math.round(paid),
            pending: Math.round(cv.outstanding),
            status: l.status,
          };
        });
        const totLoan = loanDetails.reduce((s, r) => s + r.principal, 0);
        const totInt = loanDetails.reduce((s, r) => s + r.interest, 0);
        const totPaid = loanDetails.reduce((s, r) => s + r.paid, 0);
        const totPending = loanDetails.reduce((s, r) => s + r.pending, 0);
        return { c, displayName, loanDetails, totLoan, totInt, totPaid, totPending };
      });
  }, [customers, loans, payments, selYear, searchQ]);

  const grandTotLoan = custRows.reduce((s, r) => s + r.totLoan, 0);
  const grandTotInt = custRows.reduce((s, r) => s + r.totInt, 0);
  const grandTotPaid = custRows.reduce((s, r) => s + r.totPaid, 0);
  const grandTotPending = custRows.reduce((s, r) => s + r.totPending, 0);

  const exportPDF = async () => {
    const shop = (db.settings || {}).shopName || 'Kuber Jewellery';
    let body = '';
    custRows.forEach((row) => {
      body += `<tr style="background:#EEF4FF"><td colspan="2" style="font-weight:bold">${row.displayName} (${row.c.custNo || '-'})</td><td colspan="2">${row.c.phone || '-'}</td><td style="text-align:right;font-weight:bold">${inr(row.totLoan)}</td><td style="text-align:right">${inr(row.totInt)}</td><td style="text-align:right">${inr(row.totPaid)}</td><td style="text-align:right;color:#B93A2E">${inr(row.totPending)}</td></tr>`;
      row.loanDetails.forEach((l) => {
        body += `<tr><td style="padding-left:16px">${l.loanNo}</td><td>${dmy(l.loanDate)}</td><td>${inr(l.principal)}</td><td>${inr(l.prevBalance)}</td><td style="text-align:right">${inr(l.principal)}</td><td style="text-align:right">${inr(l.interest)}</td><td style="text-align:right">${inr(l.paid)}</td><td style="text-align:right;color:#B93A2E">${inr(l.pending)}</td></tr>`;
      });
    });
    const html = `<html><head><meta charset="utf-8"><style>body{font-family:sans-serif;padding:16px;font-size:11px}h2{color:#12263A}table{width:100%;border-collapse:collapse}th{background:#12263A;color:#fff;padding:6px 8px;text-align:left;font-size:10px}td{padding:5px 8px;border-bottom:1px solid #eee}.tot{background:#D6E8FF;font-weight:bold}</style></head><body><h2>${shop}</h2><h3 style="color:#B8912F">Year-wise Member Report - ${selYear}</h3><table><thead><tr><th>Loan No</th><th>Loan Date</th><th>Principal</th><th>Prev Balance</th><th>Total Paid</th><th>Total Interest</th><th>Paid Amount</th><th>Pending</th></tr></thead><tbody>${body}</tbody><tfoot><tr class="tot"><td colspan="4">Total (${custRows.length} members)</td><td style="text-align:right">${inr(grandTotLoan)}</td><td style="text-align:right">${inr(grandTotInt)}</td><td style="text-align:right">${inr(grandTotPaid)}</td><td style="text-align:right">${inr(grandTotPending)}</td></tr></tfoot></table></body></html>`;
    if (Platform.OS === 'web') { const w = window.open('','_blank'); w.document.write(html); w.document.close(); w.print(); }
    else { const { uri } = await Print.printToFileAsync({ html }); await Sharing.shareAsync(uri, { mimeType: 'application/pdf', dialogTitle: 'Export PDF' }); }
  };

  const exportExcel = () => {
    const shop = (db.settings || {}).shopName || 'Kuber Jewellery';
    const wb = XLSX.utils.book_new();
    const hdr = ['Customer Name','Customer No','Mobile','Year','Loan Date','Principal','Interest','Prev Balance','Paid','Pending','Total Paid','Total Interest'];
    const data = [];
    custRows.forEach((row) => {
      row.loanDetails.forEach((l) => {
        data.push([row.displayName, row.c.custNo || '-', row.c.phone || '-', selYear, dmy(l.loanDate), l.principal, l.interest, l.prevBalance, l.paid, l.pending, row.totPaid, row.totInt]);
      });
    });
    const ws = XLSX.utils.aoa_to_sheet([[shop + ' - ' + selYear + ' Year-wise Member Report'],[],hdr,...data]);
    ws['!cols'] = [{wch:20},{wch:12},{wch:14},{wch:6},{wch:14},{wch:12},{wch:10},{wch:14},{wch:14},{wch:14},{wch:12},{wch:12}];
    XLSX.utils.book_append_sheet(wb, ws, selYear + ' Report');
    if (Platform.OS === 'web') { XLSX.writeFile(wb, 'yearwise_member_report_' + selYear + '.xlsx'); }
    else {
      const wbout = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
      const path = FileSystem.cacheDirectory + 'yearwise_report.xlsx';
      FileSystem.writeAsStringAsync(path, wbout, { encoding: FileSystem.EncodingType.Base64 })
        .then(() => Sharing.shareAsync(path, { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', dialogTitle: 'Export Excel' }));
    }
  };

  return (
    <View style={{ flex: 1, backgroundColor: '#fff' }}>
      <View style={{ backgroundColor: '#12263A', padding: 14 }}>
        <Text style={{ color: '#B8912F', fontWeight: '800', fontSize: 16 }}>ವರ್ಷವಾರು ಸದಸ್ಯರ ವರದಿ</Text>
        <Text style={{ color: '#B9C6D3', fontSize: 12, marginTop: 2 }}>Year-wise Member Report</Text>
      </View>
      <View style={{ backgroundColor: '#F7FAFF', borderBottomWidth: 1, borderColor: '#D6E8FF', padding: 12 }}>
        <ScrollView horizontal showsHorizontalScrollIndicator={false} style={{ marginBottom: 8 }}>
          <View style={{ flexDirection: 'row', gap: 6 }}>
            {years.map((y) => (
              <TouchableOpacity key={y} onPress={() => setSelYear(y)}
                style={{ paddingHorizontal: 14, paddingVertical: 8, borderRadius: 18, borderWidth: 1,
                  borderColor: selYear === y ? '#B8912F' : '#D0DCF0', backgroundColor: selYear === y ? '#B8912F' : '#fff' }}>
                <Text style={{ color: selYear === y ? '#fff' : '#1B2733', fontWeight: '700', fontSize: 13 }}>{y}</Text>
              </TouchableOpacity>
            ))}
          </View>
        </ScrollView>
        <View style={{ flexDirection: 'row', alignItems: 'center', backgroundColor: '#fff', borderRadius: 10, borderWidth: 1, borderColor: '#E1E6EC', paddingHorizontal: 10 }}>
          <Text style={{ color: '#6B7A8A', fontSize: 15, marginRight: 6 }}>🔍</Text>
          <TextInput value={searchQ} onChangeText={setSearchQ} placeholder="Search by name or mobile"
            placeholderTextColor="#9AA7B4" style={{ flex: 1, paddingVertical: 9, fontSize: 14, color: '#1B2733' }} />
        </View>
      </View>
      <View style={{ flexDirection: 'row', alignItems: 'center', paddingHorizontal: 12, paddingVertical: 8, backgroundColor: '#F7FAFF', borderBottomWidth: 1, borderColor: '#D6E8FF', gap: 6 }}>
        <Text style={{ flex: 1, color: '#6B7A8A', fontSize: 12 }}>{custRows.length} members · {selYear}</Text>
        <TouchableOpacity onPress={() => setViewMode(true)}
          style={{ backgroundColor: '#12263A', borderRadius: 8, paddingHorizontal: 10, paddingVertical: 6 }}>
          <Text style={{ color: '#fff', fontWeight: '700', fontSize: 12 }}>👁 View</Text>
        </TouchableOpacity>
        <TouchableOpacity onPress={exportExcel} style={{ backgroundColor: '#1E8A58', borderRadius: 8, paddingHorizontal: 10, paddingVertical: 6 }}>
          <Text style={{ color: '#fff', fontWeight: '700', fontSize: 12 }}>📥 Excel</Text>
        </TouchableOpacity>
        <TouchableOpacity onPress={exportPDF} style={{ backgroundColor: '#B93A2E', borderRadius: 8, paddingHorizontal: 10, paddingVertical: 6 }}>
          <Text style={{ color: '#fff', fontWeight: '700', fontSize: 12 }}>🖨 PDF</Text>
        </TouchableOpacity>
      </View>

      {custRows.length === 0 && (
        <View style={{ flex: 1, alignItems: 'center', justifyContent: 'center' }}>
          <Text style={{ color: '#6B7A8A', fontSize: 14 }}>No records for {selYear}.</Text>
        </View>
      )}

      {/* Full-screen View Modal */}
      <Modal visible={viewMode} animationType="slide" onRequestClose={() => setViewMode(false)}>
        <View style={{ flex: 1, backgroundColor: '#F4F7FB' }}>

          {/* Header */}
          <View style={{ backgroundColor: '#12263A', paddingTop: 16, paddingBottom: 14, paddingHorizontal: 16, flexDirection: 'row', alignItems: 'center', gap: 10 }}>
            <TouchableOpacity onPress={() => setViewMode(false)}
              style={{ backgroundColor: '#ffffff22', borderRadius: 8, paddingHorizontal: 10, paddingVertical: 6 }}>
              <Text style={{ color: '#fff', fontSize: 15, fontWeight: '700' }}>← Back</Text>
            </TouchableOpacity>
            <View style={{ flex: 1 }}>
              <Text style={{ color: '#B8912F', fontWeight: '800', fontSize: 16 }}>ವರ್ಷವಾರು ಸದಸ್ಯರ ವರದಿ</Text>
              <Text style={{ color: '#B9C6D3', fontSize: 11, marginTop: 1 }}>Year-wise Member Report · {selYear}</Text>
            </View>
            <TouchableOpacity onPress={exportExcel}
              style={{ backgroundColor: '#1E8A58', borderRadius: 8, paddingHorizontal: 10, paddingVertical: 6 }}>
              <Text style={{ color: '#fff', fontWeight: '700', fontSize: 12 }}>📥 Excel</Text>
            </TouchableOpacity>
            <TouchableOpacity onPress={exportPDF}
              style={{ backgroundColor: '#B93A2E', borderRadius: 8, paddingHorizontal: 10, paddingVertical: 6 }}>
              <Text style={{ color: '#fff', fontWeight: '700', fontSize: 12 }}>🖨 PDF</Text>
            </TouchableOpacity>
          </View>

          {/* Summary Bar */}
          <View style={{ flexDirection: 'row', backgroundColor: '#12263A', paddingHorizontal: 12, paddingBottom: 14, gap: 8 }}>
            {[['Members', String(custRows.length), '#fff'], ['Total Loan', inr(grandTotLoan), '#B8912F'], ['Paid', inr(grandTotPaid), '#1E8A58'], ['Pending', inr(grandTotPending), '#FF6B6B']].map(([lbl, val, col]) => (
              <View key={lbl} style={{ flex: 1, backgroundColor: '#ffffff11', borderRadius: 8, padding: 8, alignItems: 'center' }}>
                <Text style={{ color: col, fontWeight: '800', fontSize: 13 }} numberOfLines={1} adjustsFontSizeToFit>{val}</Text>
                <Text style={{ color: '#B9C6D3', fontSize: 10, marginTop: 2 }}>{lbl}</Text>
              </View>
            ))}
          </View>

          <ScrollView style={{ flex: 1 }} contentContainerStyle={{ padding: 12, paddingBottom: 40 }}>
            {custRows.map((row) => (
              <View key={row.c.id} style={{ marginBottom: 14, backgroundColor: '#fff', borderRadius: 14, elevation: 2, shadowColor: '#000', shadowOpacity: 0.06, shadowRadius: 6, shadowOffset: { width: 0, height: 2 }, overflow: 'hidden' }}>

                {/* Customer Header */}
                <View style={{ backgroundColor: '#12263A', paddingHorizontal: 14, paddingVertical: 12, flexDirection: 'row', justifyContent: 'space-between', alignItems: 'center' }}>
                  <View style={{ flexDirection: 'row', alignItems: 'center', gap: 10 }}>
                    <View style={{ width: 38, height: 38, borderRadius: 19, backgroundColor: '#B8912F', alignItems: 'center', justifyContent: 'center' }}>
                      <Text style={{ color: '#fff', fontWeight: '800', fontSize: 16 }}>{(row.displayName || '?')[0].toUpperCase()}</Text>
                    </View>
                    <View>
                      <Text style={{ color: '#fff', fontWeight: '800', fontSize: 15 }}>{row.displayName}</Text>
                      <Text style={{ color: '#B9C6D3', fontSize: 12, marginTop: 1 }}>{row.c.custNo || '-'} · {row.c.phone || '-'}</Text>
                    </View>
                  </View>
                  <View style={{ alignItems: 'flex-end' }}>
                    <Text style={{ color: '#B8912F', fontWeight: '800', fontSize: 15 }}>{inr(row.totLoan)}</Text>
                    <Text style={{ color: '#B9C6D3', fontSize: 10 }}>ಸಾಲದ ಮೊತ್ತ</Text>
                  </View>
                </View>

                {/* Totals Row */}
                <View style={{ flexDirection: 'row', borderBottomWidth: 1, borderColor: '#F0F4FA' }}>
                  {[['ಒಟ್ಟು ಬಡ್ಡಿ', inr(row.totInt), '#B8912F'], ['ಪಾವತಿಸಿದದು', inr(row.totPaid), '#1E8A58'], ['ಬಾಕಿ', inr(row.totPending), '#B93A2E']].map(([lbl, val, col], i) => (
                    <View key={lbl} style={{ flex: 1, paddingVertical: 10, alignItems: 'center', borderRightWidth: i < 2 ? 1 : 0, borderColor: '#F0F4FA' }}>
                      <Text style={{ color: col, fontWeight: '800', fontSize: 13 }}>{val}</Text>
                      <Text style={{ color: '#6B7A8A', fontSize: 10, marginTop: 2 }}>{lbl}</Text>
                    </View>
                  ))}
                </View>

                {/* Loan Rows */}
                {row.loanDetails.map((l, li) => (
                  <View key={l.loanNo} style={{ paddingHorizontal: 14, paddingVertical: 10, borderBottomWidth: li < row.loanDetails.length - 1 ? 1 : 0, borderColor: '#F0F4FA', backgroundColor: li % 2 === 0 ? '#fff' : '#FAFBFF' }}>
                    <View style={{ flexDirection: 'row', justifyContent: 'space-between', marginBottom: 6 }}>
                      <View style={{ flexDirection: 'row', alignItems: 'center', gap: 6 }}>
                        <View style={{ backgroundColor: '#EEF4FF', borderRadius: 6, paddingHorizontal: 8, paddingVertical: 3 }}>
                          <Text style={{ color: '#12263A', fontWeight: '700', fontSize: 12 }}>{l.loanNo}</Text>
                        </View>
                        <Text style={{ color: '#6B7A8A', fontSize: 11 }}>{dmy(l.loanDate)}</Text>
                      </View>
                      <View style={{ backgroundColor: l.status === 'active' ? '#E6F9F0' : '#F5F5F5', borderRadius: 6, paddingHorizontal: 8, paddingVertical: 3 }}>
                        <Text style={{ color: l.status === 'active' ? '#1E8A58' : '#6B7A8A', fontSize: 11, fontWeight: '600' }}>
                          {l.status === 'active' ? '● Active' : '✓ Closed'}
                        </Text>
                      </View>
                    </View>
                    <View style={{ flexDirection: 'row', flexWrap: 'wrap', gap: 8 }}>
                      {[['Principal', inr(l.principal), '#12263A'], ['Interest', inr(l.interest), '#B8912F'], ['Prev Bal', inr(l.prevBalance), '#6B7A8A'], ['Paid', inr(l.paid), '#1E8A58'], ['Pending', inr(l.pending), '#B93A2E']].map(([k, v, col]) => (
                        <View key={k} style={{ minWidth: 90, backgroundColor: '#F7FAFF', borderRadius: 8, padding: 8 }}>
                          <Text style={{ color: '#9AA7B4', fontSize: 10, marginBottom: 2 }}>{k}</Text>
                          <Text style={{ color: col, fontWeight: '700', fontSize: 13 }}>{v}</Text>
                        </View>
                      ))}
                    </View>
                  </View>
                ))}
              </View>
            ))}

            {/* Grand Total Card */}
            <View style={{ backgroundColor: '#12263A', borderRadius: 14, padding: 16, marginTop: 4 }}>
              <Text style={{ color: '#B8912F', fontWeight: '800', fontSize: 15, marginBottom: 12 }}>{selYear} · Grand Total</Text>
              {[['Total Loan Amount', inr(grandTotLoan), '#fff'], ['Total Interest', inr(grandTotInt), '#B8912F'], ['Total Paid', inr(grandTotPaid), '#1E8A58'], ['Total Pending', inr(grandTotPending), '#FF6B6B']].map(([lbl, val, col]) => (
                <View key={lbl} style={{ flexDirection: 'row', justifyContent: 'space-between', paddingVertical: 8, borderBottomWidth: 1, borderColor: '#ffffff22' }}>
                  <Text style={{ color: '#B9C6D3', fontSize: 13 }}>{lbl}</Text>
                  <Text style={{ color: col, fontWeight: '800', fontSize: 14 }}>{val}</Text>
                </View>
              ))}
            </View>
          </ScrollView>
        </View>
      </Modal>
    </View>
  );
}
