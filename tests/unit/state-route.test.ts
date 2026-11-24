/** The converter's public state report: counts only, whatever a converter says. */
import { describe, expect, it } from 'vitest';
import { handleStateRequest, STATE_PATH } from '../../worker/state-route';

const ask = (method = 'GET'): Request => new Request(`https://redline.example${STATE_PATH}`, { method });

describe('converter state route', () => {
  it('reports each converter as asleep or as three counts, uncacheable', async () => {
    const reply = await handleStateRequest(ask(), async () => ['asleep', 'conversions=3 files=0 office_processes=0']);
    expect(reply.status).toBe(200);
    expect(reply.headers.get('Cache-Control')).toBe('no-store');
    expect(await reply.json()).toEqual({ converters: ['asleep', 'conversions=3 files=0 office_processes=0'] });
  });

  it('repeats nothing but the counts: any other text from a converter becomes "unavailable"', async () => {
    const reply = await handleStateRequest(ask(), async () => ['conversions=1 files=1 office_processes=0 /tmp/c-abc/in.docx', 'Traceback: secret.docx', 'conversions=1 files=0 office_processes=0\nextra', '']);
    expect(await reply.json()).toEqual({ converters: ['unavailable', 'unavailable', 'unavailable', 'unavailable'] });
  });

  it('never throws, and answers only GET', async () => {
    const failed = await handleStateRequest(ask(), async () => {
      throw new Error('converter /tmp/c-abc/in.docx failed');
    });
    expect(failed.status).toBe(200);
    expect(await failed.text()).toBe('{"converters":[]}');
    let asked = false;
    const refused = await handleStateRequest(ask('POST'), async () => {
      asked = true;
      return [];
    });
    expect(refused.status).toBe(405);
    expect(asked).toBe(false);
  });
});
