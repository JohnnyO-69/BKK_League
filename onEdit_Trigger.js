function onEdit(e) {
  onEdit_Trigger(e);
}

function onEdit_Trigger(e) {
  if (!e) return;

  handleCurrentMatchFormEdit_(e);
  handleCurrentMatchFormV2Edit_(e);
}
