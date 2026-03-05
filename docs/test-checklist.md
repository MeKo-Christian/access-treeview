# Access Manual Test Checklist

## Prerequisites
- [ ] Windows machine with 64-bit MS Access
- [ ] Both DLLs registered via regasm
- [ ] Demo database set up per demo/README.md

## COM Registration
- [ ] `CreateObject("MeKo.TreeEngine")` succeeds in VBA Immediate window
- [ ] TreeViewHost control appears in ActiveX control list

## Form Designer
- [ ] Insert TreeViewHost control on a form
- [ ] Control renders in design view
- [ ] Form properties show control

## Basic Functionality
- [ ] Form opens and tree loads root nodes
- [ ] Expand node shows children (lazy loaded)
- [ ] Collapse node works
- [ ] Click node fires NodeClick event
- [ ] Double-click node fires NodeDoubleClick event
- [ ] Selection change fires AfterSelect event
- [ ] SelectedNodeId returns correct value

## Search
- [ ] FindAndSelect finds and selects existing node
- [ ] FindAndSelect returns False for non-existent text
- [ ] FindAndSelect expands parent chain to reach deep node

## Stability
- [ ] Close and reopen form -- control still works
- [ ] Close and reopen database -- control still works
- [ ] Multiple instances of control on different forms
- [ ] Compile database to ACCDE -- control still works
- [ ] Rapid expand/collapse -- no crashes or visual glitches

## Performance
- [ ] Tree with 100 nodes -- instant load
- [ ] Tree with 1,000 nodes -- no visible lag on expand
- [ ] Tree with 10,000+ nodes -- lazy loading keeps UI responsive
- [ ] Search on large tree -- responds within 1 second

## Error Handling
- [ ] OnError event fires when engine throws
- [ ] Missing table name -- graceful error message
- [ ] Invalid connection string -- graceful error message

## CheckBoxes
- [ ] CheckBoxes = True shows checkboxes
- [ ] CheckBoxes = False hides checkboxes
