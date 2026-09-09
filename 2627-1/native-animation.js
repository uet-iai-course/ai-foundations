/* Source-authored PPTX visibility and media steps, integrated with Reveal fragments. */
(function (global) {
  'use strict';

  function targetElements(slide, target) {
    const shapes = Array.from(slide.querySelectorAll('[data-shape-id]'))
      .filter(node => node.dataset.shapeId === String(target.shape_id));
    if (target.paragraph_start === undefined) return shapes;
    return shapes.flatMap(shape => Array.from(shape.querySelectorAll('[data-paragraph-index]')))
      .filter(node => Number(node.dataset.paragraphIndex) >= target.paragraph_start &&
        Number(node.dataset.paragraphIndex) <= target.paragraph_end);
  }

  function applyState(slide, data, index) {
    cancelTransitions(slide);
    slide.querySelectorAll('[data-native-hidden]').forEach(node => node.removeAttribute('data-native-hidden'));
    function change(target, hidden) {
      targetElements(slide, target).forEach(node => {
        if (hidden) node.setAttribute('data-native-hidden', '');
        else node.removeAttribute('data-native-hidden');
      });
    }
    function replay(events) {
      (events || []).filter(event => event.action === 'show' || event.action === 'hide')
        .forEach(event => change(event.target, event.action === 'hide'));
    }
    (data.initial_hidden || []).forEach(target => change(target, true));
    replay(data.initial_events);
    for (let i = 0; i <= index && i < data.steps.length; i++) replay(data.steps[i].events);
    slide.dataset.nativeAnimationStep = String(index);
  }

  const transitions = new WeakMap();
  let clipSequence = 0;
  function cancelTransitions(slide) {
    (transitions.get(slide) || []).forEach(cancel => cancel());
    transitions.delete(slide);
  }

  function startTransitions(slide, data, previous, index) {
    // Seeking and re-entry use the completed source state. Animate one forward
    // click only, so an interrupted exit cannot keep hiding a rewound shape.
    if (previous === null || index !== previous + 1 || index < 0) return;
    playTransitions(slide, data.steps[index].events || []);
  }

  function startInitialTransitions(slide, data, index) {
    if (index !== -1) return;
    playTransitions(slide, (data.initial_events || []).filter(event => event.transition));
  }

  function playTransitions(slide, events) {
    const cleanups = [];
    events.forEach(event => {
      if (event.action === 'show' && event.transition?.filter === 'appear') {
        const view = slide.ownerDocument.defaultView;
        const nodes = targetElements(slide, event.target);
        nodes.forEach(node => node.setAttribute('data-native-hidden', ''));
        let frame, finished = false;
        function cleanup() {
          if (finished) return;
          finished = true; view.cancelAnimationFrame(frame);
          nodes.forEach(node => node.removeAttribute('data-native-hidden'));
        }
        cleanups.push(cleanup);
        const start = view.performance.now();
        function tick(now) {
          if (finished) return;
          if (now - start >= event.transition.delay_ms) { cleanup(); return; }
          frame = view.requestAnimationFrame(tick);
        }
        frame = view.requestAnimationFrame(tick);
        return;
      }
      if (['show', 'hide'].includes(event.action) && event.transition?.filter === 'fade') {
        const view = slide.ownerDocument.defaultView;
        const entering = event.action === 'show', delay = event.transition.delay_ms || 0;
        const duration = event.transition.duration_ms;
        const terminal = event.transition.visibility_delay_ms ?? duration;
        const restorers = [], movers = [];
        targetElements(slide, event.target).forEach(node => {
          const opacity = node.style.getPropertyValue('opacity');
          const priority = node.style.getPropertyPriority('opacity');
          const base = Number(view.getComputedStyle(node).opacity);
          node.removeAttribute('data-native-hidden');
          movers.push(fraction => node.style.setProperty('opacity', String(base * fraction), priority));
          restorers.push(() => {
            if (opacity) node.style.setProperty('opacity', opacity, priority);
            else node.style.removeProperty('opacity');
            if (entering) node.removeAttribute('data-native-hidden');
            else node.setAttribute('data-native-hidden', '');
          });
        });
        let frame, finished = false;
        function cleanup() {
          if (finished) return;
          finished = true; view.cancelAnimationFrame(frame);
          restorers.forEach(restore => restore());
        }
        cleanups.push(cleanup);
        movers.forEach(move => move(entering ? 0 : 1));
        const start = view.performance.now();
        function tick(now) {
          if (finished) return;
          const elapsed = now - start - delay;
          if (elapsed >= terminal) { cleanup(); return; }
          const fraction = Math.max(0, Math.min(1, elapsed / duration));
          movers.forEach(move => move(entering ? fraction : 1 - fraction));
          frame = view.requestAnimationFrame(tick);
        }
        frame = view.requestAnimationFrame(tick);
        return;
      }
      if ((event.action === 'show' && event.transition?.filter === 'fly-in' && event.transition.direction === 'from-left') ||
          (event.action === 'hide' && event.transition?.filter === 'fly-out' && event.transition.direction === 'to-right')) {
        const document = slide.ownerDocument, view = document.defaultView;
        const stage = slide.querySelector('.native-stage');
        const sourceWidth = slide.querySelector('svg.native-layer')?.viewBox?.baseVal?.width;
        const width = stage?.clientWidth;
        const spec = event.transition, entering = event.action === 'show';
        if (!(sourceWidth > 0 && width > 0 && Number.isFinite(spec.shape_x_emu) && spec.shape_width_emu > 0)) {
          slide.dataset.nativeTransitionError = 'Horizontal fly requires native stage and source geometry';
          return;
        }
        // Left entrance starts with the right edge at x=0. Right exit ends
        // with the left edge at slideWidth; source ppt_x is a shape center.
        const distance = entering ? -spec.shape_x_emu - spec.shape_width_emu : sourceWidth - spec.shape_x_emu;
        const restorers = [], nodes = targetElements(slide, event.target);
        const movers = nodes.map(node => {
          node.removeAttribute('data-native-hidden');
          if (node.ownerSVGElement) {
            const wrapper = document.createElementNS('http://www.w3.org/2000/svg', 'g');
            node.parentNode.insertBefore(wrapper, node); wrapper.appendChild(node);
            restorers.push(() => { wrapper.parentNode.insertBefore(node, wrapper); wrapper.remove(); });
            return fraction => wrapper.setAttribute('transform', `translate(${distance * fraction} 0)`);
          }
          const transform = node.style.getPropertyValue('transform');
          const priority = node.style.getPropertyPriority('transform');
          restorers.push(() => {
            if (transform) node.style.setProperty('transform', transform, priority);
            else node.style.removeProperty('transform');
          });
          return fraction => node.style.setProperty('transform',
            `translateX(${distance * width / sourceWidth * fraction}px)${transform ? ' ' + transform : ''}`, priority);
        });
        let frame, finished = false;
        function cleanup() {
          if (finished) return;
          finished = true; view.cancelAnimationFrame(frame);
          restorers.forEach(restore => restore());
          nodes.forEach(node => {
            if (entering) node.removeAttribute('data-native-hidden');
            else node.setAttribute('data-native-hidden', '');
          });
        }
        cleanups.push(cleanup);
        movers.forEach(move => move(entering ? 1 : 0));
        const start = view.performance.now();
        function tick(now) {
          if (finished) return;
          const elapsed = now - start;
          if (elapsed >= (entering ? spec.duration_ms : spec.visibility_delay_ms)) { cleanup(); return; }
          const progress = Math.max(0, Math.min(1, elapsed / spec.duration_ms));
          movers.forEach(move => move(entering ? 1 - progress : progress));
          frame = view.requestAnimationFrame(tick);
        }
        frame = view.requestAnimationFrame(tick);
        return;
      }
      if (event.action === 'show' && event.transition?.filter === 'fly-in') {
        const document = slide.ownerDocument, view = document.defaultView;
        const stage = slide.querySelector('.native-stage');
        const layer = slide.querySelector('svg.native-layer');
        const sourceHeight = layer?.viewBox?.baseVal?.height;
        const height = stage?.clientHeight;
        const sourceY = event.transition.shape_y_emu;
        const paragraph = event.transition.target_geometry === 'paragraph';
        if (!(sourceHeight > 0 && height > 0 && (paragraph || Number.isFinite(sourceY)))) {
          slide.dataset.nativeTransitionError = 'Fly-in requires native stage and source viewBox';
          return;
        }
        // ppt_y is the normalized shape center. Source start 1 + ppt_h/2
        // places the shape top at the slide bottom: dy = slideHeight - top.
        const shapeDistance = sourceHeight - sourceY;
        const restorers = [];
        const movers = targetElements(slide, event.target).map(node => {
          // p:pRg targets a paragraph's own rectangle. Resolve its final top
          // after source text layout, including inherited placeholder geometry.
          const stageRect = paragraph ? stage.getBoundingClientRect() : null;
          const distance = paragraph
            ? (stageRect.bottom - node.getBoundingClientRect().top) * sourceHeight / stageRect.height
            : shapeDistance;
          if (node.ownerSVGElement) {
            const wrapper = document.createElementNS('http://www.w3.org/2000/svg', 'g');
            node.parentNode.insertBefore(wrapper, node);
            wrapper.appendChild(node);
            restorers.push(() => { wrapper.parentNode.insertBefore(node, wrapper); wrapper.remove(); });
            return fraction => wrapper.setAttribute('transform', `translate(0 ${distance * fraction})`);
          }
          const transform = node.style.getPropertyValue('transform');
          const priority = node.style.getPropertyPriority('transform');
          restorers.push(() => {
            if (transform) node.style.setProperty('transform', transform, priority);
            else node.style.removeProperty('transform');
          });
          return fraction => node.style.setProperty('transform',
            `translateY(${distance * height / sourceHeight * fraction}px)${transform ? ' ' + transform : ''}`, priority);
        });
        let frame, finished = false;
        function cleanup() {
          if (finished) return;
          finished = true; view.cancelAnimationFrame(frame);
          restorers.forEach(restore => restore());
        }
        cleanups.push(cleanup);
        movers.forEach(move => move(1));
        const start = view.performance.now();
        function tick(now) {
          if (finished) return;
          const progress = Math.max(0, (now - start) / event.transition.duration_ms);
          if (progress >= 1) { cleanup(); return; }
          movers.forEach(move => move(1 - progress));
          frame = view.requestAnimationFrame(tick);
        }
        frame = view.requestAnimationFrame(tick);
        return;
      }
      if (event.action !== 'hide' || event.transition?.filter !== 'dissolve') return;
      const document = slide.ownerDocument, view = document.defaultView;
      targetElements(slide, event.target).forEach(node => {
        const svg = node.ownerSVGElement;
        if (!svg || typeof node.getBBox !== 'function') {
          slide.dataset.nativeTransitionError = 'Dissolve requires an SVG target';
          return;
        }
        const box = node.getBBox();
        const ns = 'http://www.w3.org/2000/svg';
        const clip = document.createElementNS(ns, 'clipPath');
        clip.id = `native-dissolve-${++clipSequence}`;
        clip.setAttribute('clipPathUnits', 'userSpaceOnUse');
        const cells = [];
        // OOXML specifies dissolve and duration, but supplies no random seed.
        // Use a repeatable shuffled grid so every replay has the same coverage.
        const columns = 48, rows = 48;
        for (let y = 0; y < rows; y++) for (let x = 0; x < columns; x++) {
          const cell = document.createElementNS(ns, 'rect');
          cell.setAttribute('x', box.x + x * box.width / columns);
          cell.setAttribute('y', box.y + y * box.height / rows);
          cell.setAttribute('width', box.width / columns + 0.01);
          cell.setAttribute('height', box.height / rows + 0.01);
          clip.appendChild(cell); cells.push(cell);
        }
        let seed = 166;
        for (let i = cells.length - 1; i > 0; i--) {
          seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
          const j = seed % (i + 1); [cells[i], cells[j]] = [cells[j], cells[i]];
        }
        svg.appendChild(clip);
        const oldClip = node.style.getPropertyValue('clip-path');
        const priority = node.style.getPropertyPriority('clip-path');
        node.style.setProperty('clip-path', `url(#${clip.id})`);
        node.removeAttribute('data-native-hidden');
        let frame, removed = 0, finished = false;
        function cleanup() {
          if (finished) return;
          finished = true; view.cancelAnimationFrame(frame);
          clip.remove();
          if (oldClip) node.style.setProperty('clip-path', oldClip, priority);
          else node.style.removeProperty('clip-path');
          node.setAttribute('data-native-hidden', '');
        }
        cleanups.push(cleanup);
        const start = view.performance.now();
        function tick(now) {
          if (finished) return;
          if (now - start >= event.transition.visibility_delay_ms) { cleanup(); return; }
          const count = Math.floor(cells.length * Math.min(1, (now - start) / event.transition.duration_ms));
          while (removed < count) cells[removed++].remove();
          frame = view.requestAnimationFrame(tick);
        }
        frame = view.requestAnimationFrame(tick);
      });
    });
    if (cleanups.length) transitions.set(slide, cleanups);
  }

  function mediaElements(slide, target) {
    return Array.from(new Set(targetElements(slide, target).flatMap(node =>
      node.matches('video, audio') ? [node] : Array.from(node.querySelectorAll('video, audio')))));
  }

  function playMedia(media) {
    function rejected(error) {
      media.dataset.nativeMediaPlayError = String(error && (error.name || error.message) || error);
    }
    try {
      const result = media.play();
      if (result && typeof result.catch === 'function') result.catch(rejected);
    } catch (error) { rejected(error); }
  }

  function seekMedia(media, seconds) {
    const time = Number(seconds || 0);
    try { media.currentTime = Number.isFinite(time) ? Math.max(0, time) : 0; }
    catch (error) { media.dataset.nativeMediaSeekError = String(error.name || error); }
  }

  function applyMediaState(slide, data, previous, index) {
    const events = step => (step.events || []).filter(event => event.action === 'play');
    function each(event, callback) { mediaElements(slide, event.target).forEach(media => callback(media, event)); }
    function play(media, event) { seekMedia(media, event.seconds); playMedia(media); }
    if (previous === null) {
      // Re-entering a slide starts each media object at its latest active cue.
      const active = new Map();
      const firstCue = new Map();
      (data.initial_events || []).filter(event => event.action === 'play').forEach(event =>
        each(event, media => active.set(media, event)));
      data.steps.forEach((step, i) => events(step).forEach(event => each(event, media => {
        if (!firstCue.has(media)) firstCue.set(media, event);
        if (i <= index) active.set(media, event);
      })));
      firstCue.forEach((event, media) => {
        if (!active.has(media)) { media.pause(); seekMedia(media, event.seconds); }
      });
      active.forEach((event, media) => play(media, event));
    } else if (index > previous) {
      data.steps.slice(previous + 1, index + 1).forEach(step => events(step).forEach(event => each(event, play)));
    } else if (index < previous) {
      // Reverse source order so the earliest crossed cue supplies the reset time.
      data.steps.slice(index + 1, previous + 1).reverse().forEach(step =>
        events(step).reverse().forEach(event => each(event, (media, cue) => {
          media.pause(); seekMedia(media, cue.seconds);
        })));
    }
  }

  function bindMediaToggle(media) {
    let pointerPaused;
    function bodyClick(event) {
      if (event.target !== media || (event.button !== undefined && event.button !== 0)) return false;
      // Native controls are in a closed UA shadow tree. Leave their bottom strip
      // and keyboard activation alone; they retain the browser's normal behavior.
      const rect = media.getBoundingClientRect();
      const scale = media.offsetHeight ? rect.height / media.offsetHeight : 1;
      return !media.controls || (event.clientY >= rect.top && event.clientY < rect.bottom - Math.min(48 * scale, rect.height / 3));
    }
    media.addEventListener('pointerdown', event => {
      pointerPaused = bodyClick(event) ? media.paused : undefined;
    }, true);
    media.addEventListener('pointercancel', () => { pointerPaused = undefined; }, true);
    media.addEventListener('click', event => {
      if (event.detail === 0 || !bodyClick(event)) { pointerPaused = undefined; return; }
      const shouldPlay = pointerPaused === undefined ? media.paused : pointerPaused;
      pointerPaused = undefined;
      event.preventDefault();
      event.stopPropagation();
      // Some native video surfaces already toggle during dispatch. Set the
      // intended state once, instead of toggling a second time after the UA.
      Promise.resolve().then(() => {
        if (shouldPlay && media.paused) playMedia(media);
        else if (!shouldPlay && !media.paused) media.pause();
      });
    }, true);
  }

  const plugin = {
    id: 'native-animation',
    init(deck) {
      const document = deck.getRevealElement().ownerDocument;
      deck.getRevealElement().querySelectorAll('video[data-native-volume]').forEach(video => {
        video.volume = Math.max(0, Math.min(1, Number(video.dataset.nativeVolume)));
      });
      const svgAnimations = Array.from(deck.getRevealElement().querySelectorAll('svg[data-native-svg-animation]'));
      svgAnimations.forEach(svg => { svg.pauseAnimations(); svg.setCurrentTime(0); });
      // Chromium can retain SVG glyph sizes from Reveal's previous scale.
      // Recreate layout after resizing, preserving the nodes and their styles.
      let resizeFrame;
      deck.on('resize', () => {
        const view = document.defaultView;
        view.cancelAnimationFrame(resizeFrame);
        resizeFrame = view.requestAnimationFrame(() => {
          const layers = Array.from(deck.getRevealElement().querySelectorAll('.native-layer'))
            .filter(layer => layer.querySelector('text'))
            .map(layer => [layer, layer.style.getPropertyValue('display'), layer.style.getPropertyPriority('display')]);
          if (!layers.length) return;
          layers.forEach(([layer]) => layer.style.setProperty('display', 'none', 'important'));
          document.documentElement.getBoundingClientRect();
          layers.forEach(([layer, display, priority]) => {
            if (display) layer.style.setProperty('display', display, priority);
            else layer.style.removeProperty('display');
          });
        });
      });
      if (!document.getElementById('native-animation-style')) {
        const style = document.createElement('style');
        style.id = 'native-animation-style';
        style.textContent = '[data-native-hidden], [data-native-hidden] * { visibility: hidden !important; }' +
          '.native-animation-marker { position: absolute !important; width: 0 !important; height: 0 !important; overflow: hidden !important; pointer-events: none !important; }';
        document.head.appendChild(style);
      }
      const records = new Map();
      deck.getRevealElement().querySelectorAll('section[data-slide-id]').forEach(slide => {
        const script = slide.querySelector('script.native-animation-data[type="application/json"]');
        const raw = script ? script.textContent : slide.getAttribute('data-native-animation');
        if (!raw) return;
        const data = JSON.parse(raw);
        if (!Array.isArray(data.steps)) throw new Error('Invalid native animation data');
        const missing = new Set();
        const targets = [...(data.initial_hidden || []), ...(data.initial_events || []).map(e => e.target),
          ...data.steps.flatMap(step => step.events.map(e => e.target)), ...(data.media_click_toggles || [])];
        targets.forEach(target => {
          if (!targetElements(slide, target).length) missing.add(JSON.stringify(target));
        });
        if (missing.size) {
          slide.dataset.nativeAnimationMissingTargets = JSON.stringify(Array.from(missing, JSON.parse));
          console.error('Native animation targets missing', slide.dataset.slideId, Array.from(missing));
        }
        data.steps.forEach((step, index) => {
          const marker = document.createElement('span');
          marker.className = 'fragment native-animation-marker';
          marker.dataset.fragmentIndex = String(index);
          marker.setAttribute('aria-hidden', 'true');
          slide.appendChild(marker);
        });
        const toggles = new Set((data.media_click_toggles || []).flatMap(target => mediaElements(slide, target)));
        toggles.forEach(bindMediaToggle);
        records.set(slide, {data, index: null});
        applyState(slide, data, -1);
      });
      let activeSlide = null;
      function renderCurrent() {
        const slide = deck.getCurrentSlide();
        const entered = slide !== activeSlide;
        if (entered && activeSlide) {
          cancelTransitions(activeSlide);
          activeSlide.querySelectorAll('video, audio').forEach(media => media.pause());
        }
        if (entered) svgAnimations.forEach(svg => {
          svg.pauseAnimations();
          if (svg.closest('section[data-slide-id]') === slide) {
            svg.setCurrentTime(0);
            svg.unpauseAnimations();
          }
        });
        activeSlide = slide;
        const record = records.get(slide);
        if (!record) return;
        const data = record.data;
        const visible = Array.from(slide.querySelectorAll('.native-animation-marker.visible'));
        const index = visible.reduce((last, node) => Math.max(last, Number(node.dataset.fragmentIndex)), -1);
        if (entered || index !== record.index) {
          applyState(slide, data, index);
          if (entered) startInitialTransitions(slide, data, index);
          else startTransitions(slide, data, record.index, index);
        }
        applyMediaState(slide, data, entered ? null : record.index, index);
        record.index = index;
      }
      ['ready', 'slidechanged', 'fragmentshown', 'fragmenthidden'].forEach(event => deck.on(event, renderCurrent));
    },
    // Exposed for deterministic forward/backward browser verification.
    applyState,
    targetElements,
    startTransitions,
    startInitialTransitions,
    cancelTransitions,
    applyMediaState,
    bindMediaToggle
  };
  global.NativeAnimation = plugin;
  if (typeof module !== 'undefined' && module.exports) module.exports = plugin;
})(typeof window !== 'undefined' ? window : globalThis);
