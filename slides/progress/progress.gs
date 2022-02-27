/**
 * @OnlyCurrentDoc Adds progress bars to a presentation.
 */
const BAR_ID = 'PROGRESS_BAR_ID';
const BAR_HEIGHT = 10; // px

/**
 * Runs when the add-on is installed.
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen();
}

/**
 * Trigger for opening a presentation.
 * @param {object} e The onOpen event.
 */
function onOpen(e) {
  SlidesApp.getUi().createAddonMenu()
      .addItem('Show progress bar', 'createBars')
      .addItem('Hide progress bar', 'deleteBars')
      .addToUi();
}

/**
 * Create a rectangle on every slide with different bar widths.
 */
function createBars() {
  deleteBars(); // Delete any existing progress bars
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  for (let i = 0; i < slides.length; ++i) {
    const ratioComplete = (i / (slides.length - 1));
    const x = 0;
    const y = presentation.getPageHeight() - BAR_HEIGHT;
    const barWidth = presentation.getPageWidth() * ratioComplete;
    if (barWidth > 0) {
      const bar = slides[i].insertShape(SlidesApp.ShapeType.RECTANGLE, x, y,
          barWidth, BAR_HEIGHT);
      /** add dark bar example **/
      bar.getFill().setSolidFill('#0D0C27');
      bar.getBorder().setTransparent();
      bar.setLinkUrl(BAR_ID);
    }
  }
}

/**
 * Deletes all progress bar rectangles.
 */
function deleteBars() {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  for (let i = 0; i < slides.length; ++i) {
    const elements = slides[i].getPageElements();
    for (const el of elements) {
      if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE &&
        el.asShape().getLink() &&
        el.asShape().getLink().getUrl() === BAR_ID) {
        el.remove();
      }
    }
  }
}
