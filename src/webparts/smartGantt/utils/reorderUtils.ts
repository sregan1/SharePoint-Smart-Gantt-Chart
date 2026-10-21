import { ITask } from '../models';

export interface IReorderChange {
  id: number;
  updates: Pick<ITask, 'sortOrder'>;
}

const bySortOrder = (a: ITask, b: ITask): number => a.sortOrder - b.sortOrder || a.id - b.id;

/**
 * Which sibling group a task belongs to: sub-tasks are ordered among the
 * other children of the same parent, top-level tasks among the tasks of the
 * same phase. Tasks may only be reordered within their own group.
 */
function groupKey(task: ITask, ids: Set<number>): string {
  return task.parentTaskId && ids.has(task.parentTaskId)
    ? `parent:${task.parentTaskId}`
    : `phase:${task.phase || ''}`;
}

/** True when `movedId` may be dropped next to `targetId` (same sibling group). */
export function canReorderOnto(allTasks: ITask[], movedId: number, targetId: number): boolean {
  if (movedId === targetId) return false;
  const moved = allTasks.find(t => t.id === movedId);
  const target = allTasks.find(t => t.id === targetId);
  if (!moved || !target) return false;
  const ids = new Set(allTasks.map(t => t.id));
  return groupKey(moved, ids) === groupKey(target, ids);
}

/**
 * Computes the sortOrder changes for moving `movedId` just before/after
 * `targetId` among its siblings. Pass the FULL, unfiltered task list so
 * hidden tasks keep their relative order. Returns only the tasks whose
 * sortOrder actually changes (empty when the drop is a no-op or not allowed).
 *
 * The group's existing sortOrder values are reused (reassigned in the new
 * order) so the group keeps its place relative to other groups. When those
 * values contain duplicates, every task is renumbered 0..n-1 instead.
 */
export function computeReorder(
  allTasks: ITask[],
  movedId: number,
  targetId: number,
  position: 'before' | 'after'
): IReorderChange[] {
  if (!canReorderOnto(allTasks, movedId, targetId)) return [];

  const ids = new Set(allTasks.map(t => t.id));
  const moved = allTasks.find(t => t.id === movedId)!;
  const key = groupKey(moved, ids);

  const all = [...allTasks].sort(bySortOrder);
  const group = all.filter(t => groupKey(t, ids) === key);

  const next = group.filter(t => t.id !== movedId);
  const targetIndex = next.findIndex(t => t.id === targetId);
  next.splice(position === 'before' ? targetIndex : targetIndex + 1, 0, moved);

  if (next.every((t, i) => t.id === group[i].id)) return [];

  const slots = group.map(t => t.sortOrder);
  const slotsAreUnique = slots.every((v, i) => i === 0 || v > slots[i - 1]);

  const newOrder = new Map<number, number>();
  if (slotsAreUnique) {
    next.forEach((t, i) => newOrder.set(t.id, slots[i]));
  } else {
    // Duplicate values can't express an order — renumber everything, with the
    // group's members refilling the positions the group occupied.
    const arranged = [...all];
    let n = 0;
    arranged.forEach((t, i) => {
      if (groupKey(t, ids) === key) arranged[i] = next[n++];
    });
    arranged.forEach((t, i) => newOrder.set(t.id, i));
  }

  const byId = new Map(allTasks.map(t => [t.id, t] as [number, ITask]));
  const changes: IReorderChange[] = [];
  newOrder.forEach((sortOrder, id) => {
    if (byId.get(id)!.sortOrder !== sortOrder) changes.push({ id, updates: { sortOrder } });
  });
  return changes;
}
