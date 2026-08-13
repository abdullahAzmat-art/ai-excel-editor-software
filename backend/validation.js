import { z } from 'zod';

export const updateSchema = z.object({
  student: z.string().min(1, "Student name is required"),
  subject: z.string().min(1, "Subject is required"),
  marks: z.union([
    z.number().int().min(0).max(100),
    z.string().regex(/^\d+$/).transform(val => parseInt(val, 10))
  ])
});

export const llmResponseSchema = z.object({
  updates: z.array(updateSchema)
});
