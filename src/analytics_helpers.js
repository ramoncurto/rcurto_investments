function computeHistoricalPerformanceAverages(trades){
  const monthTracker={};
  const yearsSet=new Set();
  let totalInvestedEUR=0,tradesCountForRR=0,totalGrossProfitEUR=0,totalGrossLossEUR=0;
  let winningTradesCount=0,losingTradesCount=0;
  const parseNumber=v=>{
    if(v===undefined||v===null||v==='') return 0;
    let s=String(v).replace(/[€,$\s]/g,'').trim();
    const lastComma=s.lastIndexOf(',');
    const lastDot=s.lastIndexOf('.');
    if(lastComma>-1 && lastDot>-1){
      if(lastComma>lastDot){
        s=s.replace(/\./g,'');
        s=s.replace(',','.');
      }else{
        s=s.replace(/,/g,'');
      }
    }else if(lastComma>-1){
      s=s.replace(/\./g,'');
      s=s.replace(',','.');
    }else{
      s=s.replace(/,/g,'');
    }
    return parseFloat(s)||0;
  };
  trades.forEach(t=>{
    if(!t['DATE IN']) return;
    const invest=parseNumber(t['INVESTMENT (EURO)']);
    const result=parseNumber(t['RESULT']);
    const d=new Date(t['DATE IN']);
    if(!isNaN(d)){
      yearsSet.add(d.getFullYear());
      const key=`${d.getFullYear()}-${d.getMonth()}`;
      monthTracker[key]=true;
    }
    totalInvestedEUR+=invest;
    tradesCountForRR++;
    if(result>=0){winningTradesCount++; totalGrossProfitEUR+=result;}else{losingTradesCount++; totalGrossLossEUR+=Math.abs(result);} 
  });
  const years=Array.from(yearsSet).sort((a,b)=>a-b);
  const uniqueMonthsWithTrades=Object.keys(monthTracker).length;
  const historicalAvgInvestmentPerOp=tradesCountForRR>0?totalInvestedEUR/tradesCountForRR:0;
  const historicalAvgTradesPerMonth=uniqueMonthsWithTrades>0?(winningTradesCount+losingTradesCount)/uniqueMonthsWithTrades:0;
  const averageWinAmount=winningTradesCount>0?totalGrossProfitEUR/winningTradesCount:0;
  const averageLossAmount=losingTradesCount>0?totalGrossLossEUR/losingTradesCount:0;
  const winRate=(winningTradesCount+losingTradesCount)>0?winningTradesCount/(winningTradesCount+losingTradesCount):0;
  const suggestedRoiTargetPercent=(winningTradesCount>0&&historicalAvgInvestmentPerOp>0)?averageWinAmount/historicalAvgInvestmentPerOp:0;
  const suggestedLossPerFailedOpPercent=(losingTradesCount>0&&historicalAvgInvestmentPerOp>0)?averageLossAmount/historicalAvgInvestmentPerOp:0;
  return{success:true,historicalWinRate:winRate,historicalAvgInvestmentPerOp,historicalAvgTradesPerMonth,suggestedRoiTargetPercent,suggestedLossPerFailedOpPercent,years};
}
module.exports={computeHistoricalPerformanceAverages};
