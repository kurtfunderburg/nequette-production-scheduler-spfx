import { SchedulerEngine, type SchedulerData } from "./SchedulerEngine";
import { SharePointDataService, type ISharePointClient } from "./SharePointDataService";

export class ProductionSchedulerWebPart {
  private readonly schedulerEngine = new SchedulerEngine();
  private readonly dataService: SharePointDataService;

  constructor(client: ISharePointClient) {
    this.dataService = new SharePointDataService(client);
  }

  public async render(container: HTMLElement): Promise<void> {
    const schedulerData = await this.dataService.getSchedulerData();
    const normalizedData = this.schedulerEngine.normalize(schedulerData);
    const timelineRows = this.schedulerEngine.toTimelineRows(normalizedData);

    container.className = "productionScheduler";
    container.innerHTML = this.renderTable(normalizedData, timelineRows.length);
  }

  private renderTable(data: SchedulerData, rowCount: number): string {
    const taskRows = data.tasks
      .map(
        (task) => `
          <tr>
            <td>${task.id}</td>
            <td>${task.title}</td>
            <td>${task.startDate}</td>
            <td>${task.endDate}</td>
            <td>${task.percentComplete ?? 0}%</td>
          </tr>`
      )
      .join("");

    return `
      <section class="productionScheduler__container">
        <h2>Production Scheduler</h2>
        <p>Total timeline rows: ${rowCount}</p>
        <table class="productionScheduler__table">
          <thead>
            <tr>
              <th>ID</th>
              <th>Task</th>
              <th>Start</th>
              <th>End</th>
              <th>Complete</th>
            </tr>
          </thead>
          <tbody>
            ${taskRows}
          </tbody>
        </table>
      </section>
    `;
  }
}
