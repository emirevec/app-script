function getCurrentYear() {
  return new Date().getFullYear();
}

function buildExpenseReportEmailBody(expenseType, year, periods, amounts) {
  return `
    <html>
      <body>
        <p>Estimado Gastón,</p>
        <p>Ver debajo el informe de gastos de **${expenseType}** del año **${year}**:</p>
        <table border="1" style="border-collapse: collapse;">
          <thead>
            <tr>
              <th>Período</th>
              ${periods.map(period => `<th>${period}</th>`).join('')}
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>Monto Total</td>
              ${amounts.map(amount => `<td>$${amount.toFixed(2)}</td>`).join('')}
            </tr>
          </tbody>
        </table>
        <p>Saludos cordiales,</p>
      </body>
    </html>
  `;
}