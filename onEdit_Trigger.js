function onEdit_Trigger(e) {
  if (!e) return;

  const range = e.range;
  const sheet = range.getSheet();

  if (sheet.getName() === 'CurrentMatchForm' &&
      range.getRow() === 18 &&
      range.getColumn() === 1) {

    buildCurrentMatchFormSmart();
    sheet.getRange('A18').clearContent();
  }
}
