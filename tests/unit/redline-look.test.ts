/**
 * The viewer recognises each change by how the Producer formats it. The viewer
 * keeps its own copy of those colours (so the page need not load the Producer);
 * this test holds the copy to the Producer's own values.
 */
import { describe, expect, it } from 'vitest';
import { REDLINE_LOOK } from '../../src/app/redline-view';
import { CHANGE_STYLE, DELETED_COLOR, FORMAT_CHANGE_FILL, INSERTED_COLOR } from '../../src/producer/static-redline';

describe("the viewer's copy of the redline's colours", () => {
  it('matches the Producer', () => {
    expect([...REDLINE_LOOK.inserted]).toEqual([INSERTED_COLOR.standard, INSERTED_COLOR.alternative]);
    expect([...REDLINE_LOOK.deleted]).toEqual([DELETED_COLOR.standard, DELETED_COLOR.alternative]);
    expect(REDLINE_LOOK.moved).toBe(CHANGE_STYLE.movedFrom.color);
    expect(REDLINE_LOOK.moved).toBe(CHANGE_STYLE.movedTo.color);
    expect(REDLINE_LOOK.formatFill).toBe(FORMAT_CHANGE_FILL);
    // The shapes the viewer looks for: double underline, single strike, double strike, double underline.
    expect(CHANGE_STYLE).toMatchObject({ inserted: { underline: 'double' }, deleted: { strike: 'single' }, movedFrom: { strike: 'double' }, movedTo: { underline: 'double' } });
  });
});
