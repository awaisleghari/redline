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

  it('with ?detail, says where the counted files are as a count per system folder, and still repeats nothing else', async () => {
    const asked: boolean[] = [];
    const detail = new Request(`https://redline.example${STATE_PATH}?detail`);
    const good = 'before_first=4 since_first=0 where=/dev/shm:0,/tmp:4,/home/converter:0';
    const reply = await handleStateRequest(detail, async (wanted) => {
      asked.push(wanted);
      return [good, 'before_first=none since_first=0 where=/tmp:0', 'before_first=0 since_first=1 where=/tmp/c-abc/Secret Deal.docx:1', 'conversions=1 files=0 office_processes=0'];
    });
    expect(asked).toEqual([true]);
    expect(await reply.json()).toEqual({ converters: [good, 'before_first=none since_first=0 where=/tmp:0', 'unavailable', 'unavailable'] });
    await handleStateRequest(ask(), async (wanted) => {
      asked.push(wanted);
      return [];
    });
    expect(asked).toEqual([true, false]);
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
