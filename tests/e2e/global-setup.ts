/** The browser tests include the PDF step, so they need the converter container running (Docker). */
import { startContainer, stopContainer } from '../pdf/container';

export default async function globalSetup(): Promise<() => void> {
  await startContainer();
  return () => stopContainer();
}
