// ExcelDataSchema.ts
import { z } from 'zod';

export const ExcelDataSchema = z.object({
  column1: z.number(),
});
