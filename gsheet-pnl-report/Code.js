function main(){
  const report = generateExpenseReport()
  sendReportByEmail(report)
}