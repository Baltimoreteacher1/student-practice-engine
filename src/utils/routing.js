export function cleanPath(pathname) {
  const path = pathname || "/";
  return path.length > 1 ? path.replace(/\/+$/, "") : path;
}

export function parsePath(pathname) {
  const path = cleanPath(pathname);

  if (path === "/") {
    return { page: "home" };
  }

  const unitMatch = path.match(/^\/units\/([^/]+)$/);
  if (unitMatch) {
    return { page: "unit", unitSlug: decodeURIComponent(unitMatch[1]) };
  }

  const activityMatch = path.match(/^\/activities\/([^/]+)$/);
  if (activityMatch) {
    return { page: "activity", activitySlug: decodeURIComponent(activityMatch[1]) };
  }

  return { page: "notFound" };
}

export function unitPath(unitSlug) {
  return `/units/${unitSlug}`;
}

export function activityPath(activitySlug) {
  return `/activities/${activitySlug}`;
}
