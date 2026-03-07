# SlideArabi — Audit Logger Integration Guide

This document describes **exactly** where to add `AuditLogger` calls in
`rtl_transforms.py`.  **Do not modify `rtl_transforms.py` directly**; treat
this guide as the spec for that future PR.

---

## 0. Setup — constructor changes

### `MasterLayoutTransformer.__init__`

Add `audit` as an optional parameter and store it:

```python
# rtl_transforms.py — MasterLayoutTransformer.__init__
from .audit_logger import AuditLogger, classify_shape_type  # new import at top of file

def __init__(self, presentation, template_registry=None, audit: AuditLogger | None = None):
    ...
    self.audit = audit or AuditLogger()   # always non-None; caller may share one instance
```

### `SlideContentTransformer.__init__`

Same pattern:

```python
# rtl_transforms.py — SlideContentTransformer.__init__
def __init__(self, presentation, template_registry=None,
             layout_classifications=None, translations=None,
             audit: AuditLogger | None = None):
    ...
    self.audit = audit or AuditLogger()
```

> **Sharing one instance across phases:** pass the *same* `AuditLogger`
> object to both transformers so the final report covers all phases.

---

## 1. `MasterLayoutTransformer._mirror_logo_images`

**Where:** After `shape.left = new_left` on line ~423.

```python
# After: shape.left = new_left
self.audit.log_transform(
    slide_idx=0,                         # 0 = master/layout (not a content slide)
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=shape.name,
    shape_type=classify_shape_type(shape),
    transform_type='mirror_x',
    before_state={'x': original_left, 'cx': original_width},
    after_state ={'x': new_left,      'cx': original_width},
    notes='master logo mirror',
)
```

**Skip logging:**  At the `continue` that fires for an out-of-bounds result:

```python
self.audit.log_skip(
    slide_idx=0,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=shape.name,
    reason=f'logo mirror OOB: new_left={new_left}',
)
```

---

## 2. `MasterLayoutTransformer._mirror_brand_elements`

**Where:** After `shape.left = new_left` (line ~375).

```python
self.audit.log_transform(
    slide_idx=0,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=shape.name,
    shape_type='text_box',
    transform_type='mirror_x',
    before_state={'x': left, 'cx': width},
    after_state ={'x': new_left, 'cx': width},
    notes='master brand element mirror',
)
```

**Negligible-change skip** (the `abs(new_left - left) < _POSITION_TOLERANCE_EMU` branch):

```python
self.audit.log_skip(
    slide_idx=0,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=shape.name,
    reason='mirror negligible (centred shape)',
)
```

---

## 3. `MasterLayoutTransformer._mirror_layout_placeholders`

**Where:** After `shape.left = new_left` in the for-loop for individual mirrors:

```python
self.audit.log_transform(
    slide_idx=0,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=shape.name,
    shape_type='placeholder',
    transform_type='mirror_x',
    before_state={'x': left, 'cx': width},
    after_state ={'x': new_left, 'cx': width},
    notes=f'layout placeholder mirror (layout_type={layout_type})',
)
```

---

## 4. `MasterLayoutTransformer._swap_two_column_placeholders`

**Where:** After `left_ph.left = ...` / `right_ph.left = ...` in the swap block:

```python
self.audit.log_transform(
    slide_idx=0,
    shape_id=getattr(left_ph, 'shape_id', id(left_ph)),
    shape_name=left_ph.name,
    shape_type='placeholder',
    transform_type='panel_swap',
    before_state={'x': left_ph_old_left, 'cx': left_ph.width},
    after_state ={'x': new_x_left,       'cx': left_ph.width},
    notes='two-column layout swap (left placeholder)',
)
self.audit.log_transform(
    slide_idx=0,
    shape_id=getattr(right_ph, 'shape_id', id(right_ph)),
    shape_name=right_ph.name,
    shape_type='placeholder',
    transform_type='panel_swap',
    before_state={'x': right_ph_old_left, 'cx': right_ph.width},
    after_state ={'x': new_x_right,       'cx': right_ph.width},
    notes='two-column layout swap (right placeholder)',
)
```

> Capture `left_ph_old_left = left_ph.left` and `right_ph_old_left = right_ph.left`
> *before* calling `swap_positions`.

---

## 5. `SlideContentTransformer._mirror_freeform_shape`

**Where:** After `shape.left = new_left` (just before the flipH guard):

```python
self.audit.log_transform(
    slide_idx=_current_slide_number,        # pass slide_number down to this method
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=shape.name,
    shape_type=classify_shape_type(shape),
    transform_type='mirror_x',
    before_state={'x': left,    'cx': width},
    after_state ={'x': new_left,'cx': width},
)
```

**OOB skip:**

```python
self.audit.log_skip(
    slide_idx=_current_slide_number,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=shape.name,
    reason=f'mirror_freeform OOB: new_left={new_left}',
)
```

> `_mirror_freeform_shape` does not currently receive `slide_number`.
> Add it as a parameter: `def _mirror_freeform_shape(self, shape, slide_width_emu, slide_number=0)`.

---

## 6. `SlideContentTransformer._remove_local_position_override`

Three exit paths each need a log call.

### 6a. Explicit mirror (overlap risk, logo-title, or size-divergence guard)

After `shape.left = new_left` in any of the three "explicit mirror" branches:

```python
self.audit.log_transform(
    slide_idx=_slide_number,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=shape.name,
    shape_type='placeholder',
    transform_type='mirror_x',
    before_state={'x': current_left, 'cx': current_width},
    after_state ={'x': new_left,     'cx': current_width},
    notes=<branch-specific note>,   # e.g. 'overlap risk', 'size-divergence guard'
)
```

### 6b. xfrm removed (position inheritance)

After `sp_pr.remove(xfrm)`:

```python
self.audit.log_transform(
    slide_idx=_slide_number,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=shape.name,
    shape_type='placeholder',
    transform_type='position_inherit',
    before_state={'ph_idx': ph_idx, 'had_xfrm': True},
    after_state ={'ph_idx': ph_idx, 'had_xfrm': False},
    notes='xfrm removed — inheriting from RTL layout',
)
```

### 6c. No matching layout placeholder

Where the method returns `False` because `layout_ph is None`:

```python
self.audit.log_skip(
    slide_idx=_slide_number,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=shape.name,
    reason=f'placeholder idx={ph_idx} not found in layout',
)
```

> Add `slide_number=0` parameter to `_remove_local_position_override` and
> thread it through from `_transform_slide`.

---

## 7. `SlideContentTransformer._set_rtl_alignment_unconditional`

Called for every text-bearing shape.  Log once per *shape* (not per paragraph)
to avoid log explosion.

**Where:** At the end of the method, after the paragraph loop, if `changes > 0`:

```python
if changes > 0:
    # Capture the first paragraph's before/after for the representative state
    self.audit.log_transform(
        slide_idx=_slide_number,
        shape_id=getattr(shape, 'shape_id', id(shape)),
        shape_name=shape.name,
        shape_type=classify_shape_type(shape),
        transform_type='text_rtl',
        before_state={'paragraphs_modified': changes},
        after_state ={'rtl': True, 'algn': 'r'},
        notes=f'unconditional RTL pass: {changes} paragraphs',
    )
```

> Add `slide_number=0` parameter to `_set_rtl_alignment_unconditional`.

---

## 8. `SlideContentTransformer._apply_translation`

Log once per *shape* after translation, if any runs were replaced.

**Where:** At the end, after the paragraph loop, if `changes > 0`:

```python
if changes > 0:
    self.audit.log_transform(
        slide_idx=_slide_number,
        shape_id=getattr(shape, 'shape_id', id(shape)),
        shape_name=shape.name,
        shape_type=classify_shape_type(shape),
        transform_type='text_rtl',
        before_state={'lang': 'en', 'paragraphs': changes},
        after_state ={'lang': 'ar-SA', 'paragraphs': changes},
        notes='translation applied',
    )
```

---

## 9. `SlideContentTransformer._transform_chart_rtl`

Log the axis-reversal step (Step 1 in the method).

**Where:** Inside the `for cat_ax in chart_elem.iter(...)` loop, after
`orientation.set('val', 'maxMin')`:

```python
# Capture before (read before mutation) and after
self.audit.log_transform(
    slide_idx=_slide_number,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=shape.name,
    shape_type='chart',
    transform_type='axis_reversal',
    before_state={'catAx_crosses': old_crosses_val, 'orientation': 'minMax'},
    after_state ={'catAx_crosses': 'max',           'orientation': 'maxMin'},
    notes=f'ax_tag={ax_tag}',
)
```

> Read `old_crosses_val = crosses.get('val', 'autoZero')` *before* setting
> the new value.  Carry `shape` and `slide_number` into `_transform_chart_rtl`
> (already present via the outer `_transform_slide` call; pass them down).

---

## 10. `SlideContentTransformer._transform_table_rtl`

**Where:** After the tblPr RTL flag is set (end of Step 4):

```python
self.audit.log_transform(
    slide_idx=_slide_number,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=shape.name,
    shape_type='table',
    transform_type='table_rtl',
    before_state={'cols': num_cols, 'rtl': False},
    after_state ={'cols': num_cols, 'rtl': True,
                  'col_order': 'reversed'},
)
```

---

## 11. `SlideContentTransformer._reverse_directional_shape`

**Where:** After `prst_geom.set('prst', action)` **or** after `xfrm.set('flipH', ...)`:

```python
self.audit.log_transform(
    slide_idx=_slide_number,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=shape.name,
    shape_type='other',
    transform_type='directional_flip',
    before_state={'prst': prst},
    after_state ={'prst': action} if action != '_flipH'
                  else {'prst': prst, 'flipH': xfrm.get('flipH')},
    notes=f'action={action}',
)
```

---

## 12. Map-overlay exemption  (`_exempt_map_overlay_shapes`)

**Where:** Inside the `for shape in overlay_shapes` loop, after
`handled_ids.add(id(shape))`:

```python
self.audit.log_exemption(
    slide_idx=slide_number,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=getattr(shape, 'name', '?'),
    shape_type=classify_shape_type(shape),
    reason='map overlay — geographic position preserved',
)
```

---

## 13. Pre-mirror panel swap (`_pre_mirror_split_panel_swap`)

**Where:** After each `shape.left = new_left` assignment in the left_shapes and
right_shapes loops:

```python
self.audit.log_transform(
    slide_idx=slide_number,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=getattr(shape, 'name', '?'),
    shape_type=classify_shape_type(shape),
    transform_type='panel_swap',
    before_state={'x': old_left},
    after_state ={'x': new_left},
    notes='pre-mirror split-panel swap',
)
```

---

## 14. `SlideContentTransformer._should_mirror_shape` — exempt full-width shapes

**Where:** At the `return False` for full-width background shapes:

```python
self.audit.log_exemption(
    slide_idx=_slide_number,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=getattr(shape, 'name', '?'),
    shape_type=classify_shape_type(shape),
    reason='full-width background shape — no mirror',
)
```

> Add `slide_number=0` parameter to `_should_mirror_shape` to carry the
> context, or inline the exemption log at the call site in `_transform_slide`.

---

## 15. Fix 14 — Cover title anchor (`_fix_cover_title_anchor`)

**Where:** After `shape.left = new_left` (line ~2624):

```python
self.audit.log_transform(
    slide_idx=slide_number,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=getattr(shape, 'name', '?'),
    shape_type='text_box',
    transform_type='cover_anchor',
    before_state={'x': int(left), 'cx': int(width)},
    after_state ={'x': new_left,  'cx': int(width)},
    notes='cover detection — title moved to right half',
)
```

---

## 16. Fix 16 — Timeline alternation (`_reverse_timeline_alternation`)

**Where:** After `right_shape.left = ll` in the paired-swap loop:

```python
self.audit.log_transform(
    slide_idx=slide_number,
    shape_id=getattr(left_shape, 'shape_id', id(left_shape)),
    shape_name=getattr(left_shape, 'name', '?'),
    shape_type='text_box',
    transform_type='timeline_swap',
    before_state={'x': ll},
    after_state ={'x': left_shape.left},
    notes='timeline pair swap (left shape)',
)
self.audit.log_transform(
    slide_idx=slide_number,
    shape_id=getattr(right_shape, 'shape_id', id(right_shape)),
    shape_name=getattr(right_shape, 'name', '?'),
    shape_type='text_box',
    transform_type='timeline_swap',
    before_state={'x': rl},
    after_state ={'x': right_shape.left},
    notes='timeline pair swap (right shape)',
)
```

---

## 17. Fix 17 — Logo row reverse (`_reverse_logo_row_order`)

**Where:** After `shape.left = new_left` in the row-reverse loop:

```python
self.audit.log_transform(
    slide_idx=slide_number,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=getattr(shape, 'name', '?'),
    shape_type='image',
    transform_type='logo_row_reverse',
    before_state={'x': orig_l},
    after_state ={'x': new_left},
    notes=f'logo row reverse (row size={n})',
)
```

---

## 18. Fix 20 — Slide number badge (`_reposition_slide_number_badge`)

**Where:** After `shape.left = margin`:

```python
self.audit.log_transform(
    slide_idx=slide_number,
    shape_id=getattr(shape, 'shape_id', id(shape)),
    shape_name=getattr(shape, 'name', '?'),
    shape_type='text_box',
    transform_type='slide_num_badge',
    before_state={'x': l},
    after_state ={'x': margin},
    notes='slide number badge moved to top-left for RTL',
)
```

---

## 19. End-of-run reporting in `pipeline.py`

After all transformation phases complete, add:

```python
# In the pipeline's run() or transform() method, after all phases:
from slidearabi.audit_logger import AuditLogger

audit = AuditLogger()   # or retrieve the shared instance
audit.deck_name = input_path.name
audit.slide_count = len(prs.slides)
# shape_count should be accumulated during the run or counted here

audit_json_path = output_path.with_suffix('.audit.json')
audit_md_path   = output_path.with_suffix('.audit.md')

audit.to_json(str(audit_json_path))
audit.to_markdown(str(audit_md_path))
audit.print_summary()
```

---

## Summary of methods to touch

| Method | Transform type logged | Notes |
|---|---|---|
| `MasterLayoutTransformer._mirror_logo_images` | `mirror_x` | slide_idx=0 |
| `MasterLayoutTransformer._mirror_brand_elements` | `mirror_x` | slide_idx=0 |
| `MasterLayoutTransformer._mirror_layout_placeholders` | `mirror_x` | slide_idx=0 |
| `MasterLayoutTransformer._swap_two_column_placeholders` | `panel_swap` | slide_idx=0 |
| `SlideContentTransformer._mirror_freeform_shape` | `mirror_x` | add `slide_number` param |
| `SlideContentTransformer._remove_local_position_override` | `mirror_x`, `position_inherit`, `skip` | add `slide_number` param |
| `SlideContentTransformer._set_rtl_alignment_unconditional` | `text_rtl` | add `slide_number` param |
| `SlideContentTransformer._apply_translation` | `text_rtl` | add `slide_number` param |
| `SlideContentTransformer._transform_chart_rtl` | `axis_reversal` | read old value before mutation |
| `SlideContentTransformer._transform_table_rtl` | `table_rtl` | — |
| `SlideContentTransformer._reverse_directional_shape` | `directional_flip` | — |
| `SlideContentTransformer._exempt_map_overlay_shapes` | `exempt` | — |
| `SlideContentTransformer._pre_mirror_split_panel_swap` | `panel_swap` | — |
| `SlideContentTransformer._should_mirror_shape` | `exempt` | inline at call site |
| `SlideContentTransformer._fix_cover_title_anchor` | `cover_anchor` | — |
| `SlideContentTransformer._reverse_timeline_alternation` | `timeline_swap` | — |
| `SlideContentTransformer._reverse_logo_row_order` | `logo_row_reverse` | — |
| `SlideContentTransformer._reposition_slide_number_badge` | `slide_num_badge` | — |
