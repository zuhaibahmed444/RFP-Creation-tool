import { z } from "zod";

// --- Request Schema ---
export const GenerateRequestSchema = z.object({
  clientName: z.string().min(1),
  workloads: z.array(z.string().min(1)).min(1),
  deploymentType: z.enum(["software", "saas"]),
  salesRepName: z.string().min(1),
  salesRepRole: z.string().min(1),
  salesRepEmail: z.string().email(),
});
export type GenerateRequest = z.infer<typeof GenerateRequestSchema>;

// --- AI Output Contract (strict) ---
const WorkloadInScopeRowSchema = z.object({
  workloadCategory: z.string(),
  workload: z.string(),
  deploymentType: z.string(),
  location: z.string(),
});

const ComponentRowSchema = z.object({
  component: z.string(),
  quantity: z.union([z.string(), z.number()]),
  role: z.string(),
});

const HardwareSizingRowSchema = z.object({
  component: z.string(),
  cpu: z.string(),
  memory: z.string(),
  storage: z.string(),
  os: z.string(),
});

const NetworkFirewallRowSchema = z.object({
  port: z.string(),
  protocol: z.string(),
  purpose: z.string(),
});

const PrerequisiteRowSchema = z.object({
  category: z.string(),
  prerequisite: z.string(),
  customerResponsibility: z.string(),
});

const TestCaseRowSchema = z.object({
  testCase: z.string(),
  description: z.string(),
  comments: z.string(),
  result: z.string(),
});

const TimelineRowSchema = z.object({
  phase: z.string(),
  date: z.string(),
  workload: z.string(),
  task: z.string(),
});

const TestCasesByWorkloadSchema = z.object({
  workloadName: z.string(),
  rows: z.array(TestCaseRowSchema),
});

const TimelinesByWorkloadSchema = z.object({
  workloadName: z.string(),
  rows: z.array(TimelineRowSchema),
});

export const AiOutputSchema = z.object({
  customerName: z.string(),
  workloads: z.array(z.string()),
  executiveSummary: z.string(),
  objective: z.string(),
  scope: z.object({
    inScope: z.array(z.string()),
    outOfScope: z.array(z.string()),
    assumptions: z.array(z.string()),
  }),
  workloadsInScopeTable: z.array(WorkloadInScopeRowSchema),
  hardwareRequirements: z.object({
    componentsTable: z.array(ComponentRowSchema),
    hardwareSizingTable: z.array(HardwareSizingRowSchema),
    documentationLinks: z.array(z.string()),
  }),
  networkingFirewallTable: z.array(NetworkFirewallRowSchema),
  prerequisitesTable: z.array(PrerequisiteRowSchema),
  testCasesByWorkload: z.array(TestCasesByWorkloadSchema),
  timelinesByWorkload: z.array(TimelinesByWorkloadSchema),
  pocClosureAndHandover: z.string(),
});

export type AiOutput = z.infer<typeof AiOutputSchema>;
