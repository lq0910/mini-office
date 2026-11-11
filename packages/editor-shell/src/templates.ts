/** 功能：保留费用报销模板；Word 模板直接通过完整 DOCX 打开，不再生成旧正文模型；时间：2026-11-11；作者：lq。 */
import {
  createBlankWorkbook,
  type CellStyle,
  type WorkbookModel,
} from '@mini-office/document-model';
import { recalculateWorkbook } from '@mini-office/editor-sheet';

export function expenseTemplate(): WorkbookModel {
  const base = createBlankWorkbook();
  const sheet = base.sheets[0];
  sheet.name = '费用报销';
  sheet.cells = {
    '0:0': { value: '费用报销单', styleId: 'title' },
    '1:0': { value: '日期', styleId: 'header' },
    '1:1': { value: '类别', styleId: 'header' },
    '1:2': { value: '金额', styleId: 'header' },
    '1:3': { value: '备注', styleId: 'header' },
    '2:0': { value: 46313, styleId: 'date' },
    '2:1': { value: '交通费' },
    '2:2': { value: 128.5, styleId: 'money' },
    '2:3': { value: '客户拜访' },
    '3:0': { value: 46314, styleId: 'date' },
    '3:1': { value: '餐饮费' },
    '3:2': { value: 380, styleId: 'money' },
    '3:3': { value: '项目沟通' },
    '4:0': { value: 46315, styleId: 'date' },
    '4:1': { value: '办公用品' },
    '4:2': { value: 259.9, styleId: 'money' },
    '4:3': { value: '打印纸、墨盒' },
    '5:0': { value: 46316, styleId: 'date' },
    '5:1': { value: '差旅费' },
    '5:2': { value: 1200, styleId: 'money' },
    '5:3': { value: '高铁票' },
    '6:0': { value: '合计', styleId: 'total' },
    '6:2': { value: null, formula: '=SUM(C3:C6)', styleId: 'totalMoney' },
    '6:3': { value: '由财务审核后报销', styleId: 'total' },
    '8:0': { value: '说明：请附上原始发票，填写完成后提交财务。' },
  };
  sheet.merges = [{ r1: 0, c1: 0, r2: 0, c2: 3 }];
  sheet.freeze = { rows: 2, cols: 0 };
  sheet.colWidths = { 0: 100, 1: 100, 2: 110, 3: 200 };
  sheet.rowHeights = { 0: 34, 1: 24 };
  sheet.autoFilter = { r1: 1, c1: 0, r2: 5, c2: 3 };
  const styles: Record<string, CellStyle> = {
    default: { fontFamily: '等线', fontSizePt: 11, numFmt: 'General', verticalAlign: 'middle' },
    title: {
      fontFamily: '微软雅黑',
      fontSizePt: 18,
      bold: true,
      horizontalAlign: 'center',
      verticalAlign: 'middle',
      color: '#1D2939',
    },
    header: {
      fontFamily: '等线',
      fontSizePt: 11,
      bold: true,
      fill: '#3478F6',
      color: '#FFFFFF',
      horizontalAlign: 'center',
      verticalAlign: 'middle',
      borders: {
        bottom: { style: 'thin', color: '#2F6EE0' },
      },
    },
    date: { fontFamily: '等线', fontSizePt: 11, numFmt: 'yyyy-mm-dd', horizontalAlign: 'center' },
    money: { fontFamily: '等线', fontSizePt: 11, numFmt: '¥#,##0.00', horizontalAlign: 'right' },
    total: {
      fontFamily: '等线',
      fontSizePt: 11,
      bold: true,
      fill: '#EEF2F9',
      horizontalAlign: 'center',
      verticalAlign: 'middle',
    },
    totalMoney: {
      fontFamily: '等线',
      fontSizePt: 12,
      bold: true,
      numFmt: '¥#,##0.00',
      fill: '#EEF2F9',
      color: '#D92D20',
      horizontalAlign: 'right',
    },
  };
  const model: WorkbookModel = {
    ...base,
    sheets: [sheet],
    activeSheetId: sheet.id,
    styles,
    defaultFont: '等线',
    defaultFontSizePt: 11,
    meta: { title: '费用报销单' },
    stats: { cells: Object.keys(sheet.cells).length, sheets: 1 },
  };
  return recalculateWorkbook(model);
}
