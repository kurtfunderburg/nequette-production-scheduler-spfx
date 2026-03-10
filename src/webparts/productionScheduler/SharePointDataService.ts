import type { SchedulerData, SchedulerMilestone, SchedulerTask } from "./SchedulerEngine";

export interface ISharePointClient {
  getListItems(listTitle: string): Promise<Record<string, unknown>[]>;
  createListItem(listTitle: string, payload: Record<string, unknown>): Promise<void>;
}

export class SharePointDataService {
  private readonly taskListName = "Scheduler Tasks";
  private readonly milestoneListName = "Scheduler Milestones";

  constructor(private readonly client: ISharePointClient) {}

  public async getSchedulerData(): Promise<SchedulerData> {
    const [taskItems, milestoneItems] = await Promise.all([
      this.client.getListItems(this.taskListName),
      this.client.getListItems(this.milestoneListName)
    ]);

    return {
      tasks: taskItems.map(this.mapTaskItem),
      milestones: milestoneItems.map(this.mapMilestoneItem)
    };
  }

  public async saveTask(task: SchedulerTask): Promise<void> {
    await this.client.createListItem(this.taskListName, {
      Title: task.title,
      SchedulerTaskId: task.id,
      StartDate: task.startDate,
      EndDate: task.endDate,
      PercentComplete: task.percentComplete ?? 0,
      Status: task.status ?? "Not Started",
      DependsOnTaskIds: JSON.stringify(task.dependsOnTaskIds ?? [])
    });
  }

  private mapTaskItem = (item: Record<string, unknown>): SchedulerTask => ({
    id: Number(item.SchedulerTaskId ?? item.ID ?? 0),
    title: String(item.Title ?? "Untitled task"),
    startDate: String(item.StartDate ?? ""),
    endDate: String(item.EndDate ?? ""),
    status: String(item.Status ?? "Not Started"),
    percentComplete: Number(item.PercentComplete ?? 0),
    dependsOnTaskIds: this.parseDependencyIds(item.DependsOnTaskIds)
  });

  private mapMilestoneItem = (item: Record<string, unknown>): SchedulerMilestone => ({
    id: Number(item.SchedulerMilestoneId ?? item.ID ?? 0),
    title: String(item.Title ?? "Untitled milestone"),
    dueDate: String(item.DueDate ?? ""),
    dependsOnTaskIds: this.parseDependencyIds(item.DependsOnTaskIds)
  });

  private parseDependencyIds(rawValue: unknown): number[] {
    if (typeof rawValue !== "string" || !rawValue.trim()) {
      return [];
    }

    try {
      const parsed = JSON.parse(rawValue);
      return Array.isArray(parsed) ? parsed.map(Number).filter((value) => !Number.isNaN(value)) : [];
    } catch {
      return [];
    }
  }
}
