export type SurfaceVector = { x: number; y: number; z: number };

export function surfaceNormal(a: SurfaceVector, b: SurfaceVector, c: SurfaceVector): SurfaceVector {
  const ux = b.x - a.x;
  const uy = b.y - a.y;
  const uz = b.z - a.z;
  const vx = c.x - a.x;
  const vy = c.y - a.y;
  const vz = c.z - a.z;
  const x = uy * vz - uz * vy;
  const y = uz * vx - ux * vz;
  const z = ux * vy - uy * vx;
  const length = Math.hypot(x, y, z) || 1;
  return { x: x / length, y: y / length, z: z / length };
}

export function transformSurfaceNormal(normal: SurfaceVector, rotX: number, rotY: number): SurfaceVector {
  const xAngle = rotX * Math.PI / 180;
  const yAngle = rotY * Math.PI / 180;
  const cx = Math.cos(xAngle);
  const sx = Math.sin(xAngle);
  const cy = Math.cos(yAngle);
  const sy = Math.sin(yAngle);
  const x = -cy * normal.x + sy * normal.z;
  const y = sy * sx * normal.x + cx * normal.y + cy * sx * normal.z;
  const z = sy * cx * normal.x - sx * normal.y + cy * cx * normal.z;
  const length = (Math.hypot(x, y, z) || 1) * (z > 0 ? -1 : 1);
  return { x: x / length, y: y / length, z: z / length };
}

export function shadeSurfaceMatteColor(color: string, normal: SurfaceVector) {
  const hex = color.replace(/^#/, "");
  if (!/^[0-9a-f]{6}$/i.test(hex)) return color;
  const diffuse = Math.max(0, Math.min(1, (normal.x + normal.y - normal.z) / Math.sqrt(3)));
  const intensity = 0.4 + 0.75 * diffuse;
  return `#${[0, 2, 4].map((offset) => Math.min(255, Math.floor(Number.parseInt(hex.slice(offset, offset + 2), 16) * intensity)).toString(16).padStart(2, "0")).join("")}`;
}
