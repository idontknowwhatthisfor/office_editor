import { drawChart } from './draw';
import { readChart, type ChartTheme } from './read';

export type { ChartTheme } from './read';

/**
 * A chart part (chartN.xml or chartExN.xml) drawn as SVG markup of `w` × `h` CSS px, or null if
 * it holds nothing this code can draw (the host then shows its placeholder).
 */
export function renderChart(root: Element, theme: ChartTheme, w: number, h: number): string | null {
  try {
    const chart = readChart(root, theme);
    return chart ? drawChart(chart, w, h) : null;
  } catch (e) {
    console.error('chart failed', e);
    return null;
  }
}
