export interface SchedulerTask {
  id: number;
  title: string;
  startDate: string;
  endDate: string;
  status?: string;
  percentComplete?: number;
  dependsOnTaskIds?: number[];
}

export interface SchedulerMilestone {
  id: number;
  title: string;
  dueDate: string;
  dependsOnTaskIds?: number[];
}

export interface SchedulerData {
  tasks: SchedulerTask[];
  milestones: SchedulerMilestone[];
}

export class SchedulerEngine {
  public normalize(data: SchedulerData): SchedulerData {
    const tasks = data.tasks.map((task) => ({
      ...task,
      percentComplete: this.clampPercent(task.percentComplete)
    }));

    return {
      tasks,
      milestones: data.milestones
    };
  }

  public toTimelineRows(data: SchedulerData): Array<Record<string, unknown>> {
    const taskRows = data.tasks.map((task) => ({
      type: "task",
      id: task.id,
      title: task.title,
      startDate: task.startDate,
      endDate: task.endDate,
      percentComplete: this.clampPercent(task.percentComplete),
      status: task.status ?? "Not Started",
      dependsOnTaskIds: task.dependsOnTaskIds ?? []
    }));

    const milestoneRows = data.milestones.map((milestone) => ({
      type: "milestone",
      id: milestone.id,
      title: milestone.title,
      dueDate: milestone.dueDate,
      dependsOnTaskIds: milestone.dependsOnTaskIds ?? []
    }));

    return [...taskRows, ...milestoneRows];
  }

  private clampPercent(value?: number): number {
    if (typeof value !== "number" || Number.isNaN(value)) {
      return 0;
    }

    return Math.min(100, Math.max(0, value));
  }
}
