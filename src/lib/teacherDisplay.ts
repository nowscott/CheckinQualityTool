export function displayTeacherName(value: unknown, email: unknown) {
  const name = String(value ?? "");
  const localPart = String(email ?? "").split("@")[0];
  const emailDigits = localPart.match(/(\d+)$/u)?.[1] || "";
  if (!emailDigits) return name;
  return `${name.replace(/[0-9０-９]+$/u, "")}${emailDigits}`;
}
