const fs=require('fs');
const {parse}=require('csv-parse/sync');
const {computeHistoricalPerformanceAverages}=require('../src/analytics_helpers');

test('computeHistoricalPerformanceAverages on sample data',()=>{
  const csv=fs.readFileSync('data/Trades.csv','utf8');
  const trades=parse(csv,{columns:true,skip_empty_lines:true}).filter(r=>r['DATE IN']);
  const res=computeHistoricalPerformanceAverages(trades);
  expect(res.success).toBe(true);
  expect(res.historicalWinRate).toBeCloseTo(0.696969,5);
  expect(res.historicalAvgInvestmentPerOp).toBeCloseTo(979.4776,1);
  expect(res.historicalAvgTradesPerMonth).toBeCloseTo(5.5,5);
  expect(res.suggestedRoiTargetPercent).toBeCloseTo(0.1739,3);
  expect(res.suggestedLossPerFailedOpPercent).toBeCloseTo(0.1493,3);
});
