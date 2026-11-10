/**
 * Map over items with at most `limit` tasks in flight. Results keep the input
 * order no matter which task finishes first.
 */
export async function mapWithConcurrency<T, R>(
  items: readonly T[],
  limit: number,
  task: (item: T, index: number) => Promise<R>,
): Promise<R[]> {
  const results = new Array<R>(items.length);
  let next = 0;
  const workerCount = Math.max(1, Math.min(limit, items.length));

  async function worker(): Promise<void> {
    while (next < items.length) {
      const index = next++;
      results[index] = await task(items[index] as T, index);
    }
  }

  await Promise.all(Array.from({ length: workerCount }, worker));
  return results;
}
