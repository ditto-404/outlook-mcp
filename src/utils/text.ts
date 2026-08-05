/** Outlook 이 변환해 준 본문 텍스트의 과도한 공백/빈 줄을 정리합니다. */
export function cleanupPlainText(input: string): string {
  return input
    .replace(/\r\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .split('\n')
    .map((line) => line.replace(/[ \t]+$/g, ''))
    .join('\n')
    .trim();
}
