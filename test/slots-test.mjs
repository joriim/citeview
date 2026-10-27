// Tests which citation goes to the left and right pane.
import assert from 'node:assert/strict';
import { assignSlots, isSlotColor } from '../src/slots.js';

const C = (rels) => rels.map((rel, i) => ({ key: 'c' + i, rel }));
const t = (name, cands, pins, exp) => {
  const got = assignSlots(cands, pins);
  assert.deepEqual(got, exp, `${name}: got ${JSON.stringify(got)}`);
  console.log('ok  ', name, JSON.stringify(got));
};

t('cursor inside 2nd citation', C(['before', 'at', 'after', 'after']), {}, { left: 1, right: 2 });
t('cursor in text before c1', C(['before', 'after', 'after']), {}, { left: 1, right: 2 });
t('cursor after last citation', C(['before', 'before']), {}, { left: 1, right: 0 });
t('only one citation', C(['after']), {}, { left: 0, right: -1 });
t('no citations', [], {}, { left: -1, right: -1 });
t('left pinned to c0, cursor at c2', C(['before', 'before', 'at', 'after']), { left: 'c0' }, { left: 0, right: 2 });
t('right pinned to c1, cursor at c1', C(['before', 'at', 'after']), { right: 'c1' }, { left: 2, right: 1 });
t('left pinned outside window', C(['at', 'after']), { left: 'zz' }, { left: -1, right: 0 });
assert.ok(isSlotColor('#00ffff', 'left') && isSlotColor('Turquoise', 'left') && isSlotColor('BrightGreen', 'right') && isSlotColor('#00ff00', 'right'));
assert.ok(!isSlotColor('#00FFFF', 'right') && !isSlotColor(null, 'left') && !isSlotColor('Yellow', 'left'));
console.log('colour checks ok');
