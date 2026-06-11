import { addDays } from 'date-fns';
import { ITask, ITaskFilter, isFilterActive } from '../models';
import { parseDateOnly, todayLocalMidnight } from './dateUtils';

export function filterTasks(tasks: ITask[], filter: ITaskFilter): ITask[] {
  if (!isFilterActive(filter)) return tasks;

  const text = filter.text.trim().toLowerCase();
  const today = todayLocalMidnight();
  const weekEnd = addDays(today, 7);

  return tasks.filter(t => {
    if (text) {
      const haystack = `${t.title} ${t.description} ${t.notes} ${t.assignedTo} ${t.phase}`.toLowerCase();
      if (haystack.indexOf(text) === -1) return false;
    }
    if (filter.statuses.length > 0 && filter.statuses.indexOf(t.status) === -1) return false;
    if (filter.priorities.length > 0 && filter.priorities.indexOf(t.priority) === -1) return false;
    if (filter.assignees.length > 0 && filter.assignees.indexOf(t.assignedTo) === -1) return false;
    if (filter.phases.length > 0 && filter.phases.indexOf(t.phase) === -1) return false;

    if (filter.due !== 'all') {
      if (t.status === 'Completed' || t.status === 'Cancelled') return false;
      const due = parseDateOnly(t.dueDate);
      if (!due) return false;
      if (filter.due === 'overdue' && !(due < today)) return false;
      if (filter.due === 'today' && due.getTime() !== today.getTime()) return false;
      if (filter.due === 'week' && !(due >= today && due <= weekEnd)) return false;
    }
    return true;
  });
}
