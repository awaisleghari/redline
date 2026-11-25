/** The landing page: two animated bitmap grids behind the headline. */
import { startGrid, type GridHandle } from './grid';

const handles: GridHandle[] = [];

function start(): void {
  if (handles.length) return;
  const top = document.getElementById('grid-top') as HTMLCanvasElement | null;
  const bottom = document.getElementById('grid-bottom') as HTMLCanvasElement | null;
  if (top) handles.push(startGrid(top, false)); /* brightest at the top, fades downward */
  if (bottom) handles.push(startGrid(bottom, true)); /* mirrored: brightest at the bottom, fades upward */
}

start();
// Leave no animation loop behind when the page goes away; start again if the browser brings the page back.
window.addEventListener('pagehide', () => handles.splice(0).forEach((handle) => handle.stop()));
window.addEventListener('pageshow', start);
