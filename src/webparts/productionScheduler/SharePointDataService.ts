import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IDeliverable {
  id: string;
  spId?: number;
  name: string;
  duration: number;
  startOffset: number;
}

export interface IMilestone {
  id: string;
  spId?: number;
  name: string;
  startOffset: number;
  duration: number;
  color: string;
  deliverables: IDeliverable[];
}

export interface IPhase {
  id: string;
  spId?: number;
  name: string;
  startDate: string;
  duration: number;
  visible: boolean;
  milestones: IMilestone[];
}

export type ViewMode = 'gantt' | 'calendar';

export interface ISchedulerState {
  projectSpId?: number;
  projectName: string;
  projectNumber: string;
  globalStartDate: string;
  viewMode: ViewMode;
  lastUpdated: number;
  phases: IPhase[];
}

interface IProjectListItem {
  Id: number;
  Title: string;
  ProjectNumber?: string;
  GlobalStartDate?: string;
  ViewMode?: ViewMode;
  LastUpdated?: number;
}

export class SharePointDataService {
  private readonly baseUrl: string;

  public constructor(
    siteUrl: string,
    private readonly spHttpClient: SPHttpClient,
    private readonly projectName: string
  ) {
    this.baseUrl = `${siteUrl}/_api/web/lists`;
  }

  public async loadData(): Promise<ISchedulerState> {
    const fallback: ISchedulerState = {
      projectName: this.projectName,
      projectNumber: '',
      globalStartDate: this.todayISO(),
      viewMode: 'gantt',
      lastUpdated: Date.now(),
      phases: []
    };

    const projects = await this.getItems<IProjectListItem>('SchedulerProjects', `?$filter=Title eq '${this.escapeSingleQuotes(this.projectName)}'&$top=1`);
    if (projects.length < 1) {
      return fallback;
    }

    const project = projects[0];
    const phases = await this.getItems<any>('SchedulerPhases', `?$filter=ProjectId eq ${project.Id}&$orderby=SortOrder asc,Id asc`);
    const milestones = await this.getItems<any>('SchedulerMilestones', `?$orderby=SortOrder asc,Id asc`);
    const deliverables = await this.getItems<any>('SchedulerDeliverables', `?$orderby=SortOrder asc,Id asc`);

    const phaseMap: Record<number, IPhase> = {};
    phases.forEach((phase: any): void => {
      phaseMap[phase.Id] = {
        id: phase.PhaseKey || `phase_${phase.Id}`,
        spId: phase.Id,
        name: phase.Title,
        startDate: this.dateOnly(phase.StartDate),
        duration: Number(phase.Duration || 1),
        visible: phase.Visible !== false,
        milestones: []
      };
    });

    const milestoneMap: Record<number, IMilestone> = {};
    milestones.forEach((ms: any): void => {
      const parentPhase = phaseMap[ms.PhaseId];
      if (!parentPhase) {
        return;
      }
      const mapped: IMilestone = {
        id: ms.MilestoneKey || `milestone_${ms.Id}`,
        spId: ms.Id,
        name: ms.Title,
        startOffset: Number(ms.StartOffset || 0),
        duration: Number(ms.Duration || 1),
        color: ms.Color || '#3b82f6',
        deliverables: []
      };
      milestoneMap[ms.Id] = mapped;
      parentPhase.milestones.push(mapped);
    });

    deliverables.forEach((d: any): void => {
      const parentMs = milestoneMap[d.MilestoneId];
      if (!parentMs) {
        return;
      }
      parentMs.deliverables.push({
        id: `deliverable_${d.Id}`,
        spId: d.Id,
        name: d.Title,
        duration: Number(d.Duration || 1),
        startOffset: Number(d.StartOffset || 0)
      });
    });

    return {
      projectSpId: project.Id,
      projectName: project.Title,
      projectNumber: project.ProjectNumber || '',
      globalStartDate: this.dateOnly(project.GlobalStartDate) || fallback.globalStartDate,
      viewMode: project.ViewMode || 'gantt',
      lastUpdated: Number(project.LastUpdated || Date.now()),
      phases: Object.keys(phaseMap).map((id: string): IPhase => phaseMap[Number(id)])
    };
  }

  public async saveData(state: ISchedulerState): Promise<void> {
    state.lastUpdated = Date.now();
    const projectId = await this.upsertProject(state);
    await this.replaceChildren(projectId, state);
  }

  public async createPhase(projectId: number, phase: IPhase, sortOrder: number): Promise<number> {
    const payload = {
      Title: phase.name,
      PhaseKey: phase.id,
      ProjectId: projectId,
      StartDate: phase.startDate,
      Duration: phase.duration,
      Visible: phase.visible,
      SortOrder: sortOrder
    };
    const created = await this.postItem('SchedulerPhases', payload);
    return created.Id;
  }

  public async createMilestone(phaseSpId: number, milestone: IMilestone, sortOrder: number): Promise<number> {
    const payload = {
      Title: milestone.name,
      MilestoneKey: milestone.id,
      PhaseId: phaseSpId,
      StartOffset: milestone.startOffset,
      Duration: milestone.duration,
      Color: milestone.color,
      SortOrder: sortOrder
    };
    const created = await this.postItem('SchedulerMilestones', payload);
    return created.Id;
  }

  public async createDeliverable(milestoneSpId: number, deliverable: IDeliverable, sortOrder: number): Promise<number> {
    const payload = {
      Title: deliverable.name,
      MilestoneId: milestoneSpId,
      StartOffset: deliverable.startOffset,
      Duration: deliverable.duration,
      SortOrder: sortOrder
    };
    const created = await this.postItem('SchedulerDeliverables', payload);
    return created.Id;
  }

  private async upsertProject(state: ISchedulerState): Promise<number> {
    const payload = {
      Title: state.projectName,
      ProjectNumber: state.projectNumber,
      GlobalStartDate: state.globalStartDate,
      ViewMode: state.viewMode,
      LastUpdated: state.lastUpdated
    };

    if (state.projectSpId) {
      await this.updateItem('SchedulerProjects', state.projectSpId, payload);
      return state.projectSpId;
    }

    const existing = await this.getItems<IProjectListItem>('SchedulerProjects', `?$filter=Title eq '${this.escapeSingleQuotes(state.projectName)}'&$top=1`);
    if (existing.length > 0) {
      state.projectSpId = existing[0].Id;
      await this.updateItem('SchedulerProjects', existing[0].Id, payload);
      return existing[0].Id;
    }

    const created = await this.postItem('SchedulerProjects', payload);
    state.projectSpId = created.Id;
    return created.Id;
  }

  private async replaceChildren(projectId: number, state: ISchedulerState): Promise<void> {
    const existingPhases = await this.getItems<{ Id: number }>('SchedulerPhases', `?$select=Id&$filter=ProjectId eq ${projectId}`);
    const phaseIds = existingPhases.map((phase: { Id: number }): number => phase.Id);

    if (phaseIds.length > 0) {
      const existingMilestones = await this.getItems<{ Id: number; PhaseId: number }>('SchedulerMilestones', `?$select=Id,PhaseId`);
      const milestoneIds = existingMilestones
        .filter((milestone: { Id: number; PhaseId: number }): boolean => phaseIds.indexOf(milestone.PhaseId) >= 0)
        .map((milestone: { Id: number }): number => milestone.Id);

      if (milestoneIds.length > 0) {
        const existingDeliverables = await this.getItems<{ Id: number; MilestoneId: number }>('SchedulerDeliverables', `?$select=Id,MilestoneId`);
        for (const deliverable of existingDeliverables.filter((item: { Id: number; MilestoneId: number }): boolean => milestoneIds.indexOf(item.MilestoneId) >= 0)) {
          await this.deleteItem('SchedulerDeliverables', deliverable.Id);
        }
      }

      for (const milestone of existingMilestones.filter((item: { Id: number; PhaseId: number }): boolean => phaseIds.indexOf(item.PhaseId) >= 0)) {
        await this.deleteItem('SchedulerMilestones', milestone.Id);
      }

      for (const phase of existingPhases) {
        await this.deleteItem('SchedulerPhases', phase.Id);
      }
    }

    for (let pIndex = 0; pIndex < state.phases.length; pIndex++) {
      const phase = state.phases[pIndex];
      const phaseId = await this.createPhase(projectId, phase, pIndex);
      phase.spId = phaseId;

      for (let mIndex = 0; mIndex < phase.milestones.length; mIndex++) {
        const milestone = phase.milestones[mIndex];
        const milestoneId = await this.createMilestone(phaseId, milestone, mIndex);
        milestone.spId = milestoneId;

        for (let dIndex = 0; dIndex < milestone.deliverables.length; dIndex++) {
          const deliverable = milestone.deliverables[dIndex];
          const deliverableId = await this.createDeliverable(milestoneId, deliverable, dIndex);
          deliverable.spId = deliverableId;
        }
      }
    }
  }

  private async getItems<T>(listTitle: string, query: string): Promise<T[]> {
    const response = await this.spHttpClient.get(
      `${this.baseUrl}/GetByTitle('${listTitle}')/items${query}`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );
    const json = await response.json();
    return (json.value || []) as T[];
  }

  private async postItem(listTitle: string, payload: unknown): Promise<any> {
    const response = await this.spHttpClient.post(
      `${this.baseUrl}/GetByTitle('${listTitle}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify(payload)
      }
    );
    return this.handleWriteResponse(response);
  }

  private async updateItem(listTitle: string, id: number, payload: unknown): Promise<void> {
    const response = await this.spHttpClient.post(
      `${this.baseUrl}/GetByTitle('${listTitle}')/items(${id})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify(payload)
      }
    );
    await this.handleWriteResponse(response);
  }

  private async deleteItem(listTitle: string, id: number): Promise<void> {
    const response = await this.spHttpClient.post(
      `${this.baseUrl}/GetByTitle('${listTitle}')/items(${id})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE'
        }
      }
    );
    await this.handleWriteResponse(response);
  }

  private async handleWriteResponse(response: SPHttpClientResponse): Promise<any> {
    if (!response.ok) {
      const text = await response.text();
      throw new Error(`SharePoint write failed (${response.status}): ${text}`);
    }

    const contentType = response.headers.get('content-type') || '';
    if (contentType.toLowerCase().indexOf('application/json') >= 0) {
      return response.json();
    }

    return undefined;
  }

  private dateOnly(value: string | undefined): string {
    if (!value) {
      return '';
    }
    return value.length >= 10 ? value.slice(0, 10) : value;
  }

  private todayISO(): string {
    return new Date().toISOString().slice(0, 10);
  }

  private escapeSingleQuotes(value: string): string {
    return value.replace(/'/g, "''");
  }
}
