/** Создаёт новый файл. Уже существующие презентации не изменяет. */
function createBookCampaignPresentation() {
  const deck = SlidesApp.create('Книжная кампания — решение о запуске');
  const width = deck.getPageWidth(), height = deck.getPageHeight();
  const first = deck.getSlides()[0];
  const data = [
    ['Решение о запуске', 'Согласовать бюджет 1,2 млн ₽\nСтарт продвижения: 15 октября 2026\nУчебный пример'],
    ['Продажи по каналам', 'Магазины: 2 400 экземпляров\nМаркетплейсы: 3 600 экземпляров\nЦифровой: 4 000 экземпляров\nИтого: 10 000 экземпляров'],
    ['Бюджет кампании', 'Реклама — 600 тыс. ₽\nКонтент — 300 тыс. ₽\nМероприятия — 180 тыс. ₽\nРезерв — 120 тыс. ₽\nИтого — 1 200 тыс. ₽'],
    ['Следующий шаг', 'Редактор согласует содержание.\nМаркетолог подтверждает календарь.\nАналитик сверяет показатели.\nЧерез неделю после запуска — отчёт.']
  ];
  data.forEach(function(item, index) {
    const slide = deck.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    slide.getBackground().setSolidFill('#1e2027');
    const title = slide.insertTextBox(item[0], width * .06, height * .08, width * .88, height * .18);
    title.getText().getTextStyle().setFontFamily('Arial').setFontSize(28).setBold(true).setForegroundColor('#ff9f83');
    const body = slide.insertTextBox(item[1], width * .06, height * .31, width * .88, height * .52);
    body.getText().getTextStyle().setFontFamily('Arial').setFontSize(21).setForegroundColor('#f5f2ec');
    const number = slide.insertTextBox(String(index + 1), width * .9, height * .9, width * .05, height * .06);
    number.getText().getTextStyle().setFontSize(10).setForegroundColor('#adb3c0');
  });
  if (first) first.remove();
  const url = deck.getUrl();
  Logger.log('Открыть презентацию: ' + url);
  return url;
}
