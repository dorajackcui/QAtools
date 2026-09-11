// Optional scroll position indicator; anchors and screenshots work without JS.
(() => {
  const sections = [...document.querySelectorAll('.guide-tool[data-tool]')];
  const links = [...document.querySelectorAll('#tool-navigation a[data-tool]')];
  let pending = false;
  function update() {
    pending = false;
    let current = sections[0];
    for (const section of sections) {
      if (section.getBoundingClientRect().top > 150) break;
      current = section;
    }
    for (const link of links) {
      if (link.dataset.tool === current?.dataset.tool) link.setAttribute('aria-current', 'location');
      else link.removeAttribute('aria-current');
    }
  }
  function schedule() {
    if (!pending) {
      pending = true;
      requestAnimationFrame(update);
    }
  }
  window.addEventListener('scroll', schedule, { passive: true });
  window.addEventListener('resize', schedule);
  window.addEventListener('hashchange', schedule);
  update();
})();
