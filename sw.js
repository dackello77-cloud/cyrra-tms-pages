self.addEventListener("push", (event) => {
  let payload = {};
  try {
    payload = event.data ? event.data.json() : {};
  } catch {
    payload = {};
  }
  const title = payload.title || "Cyrra TMS";
  const actionUrl = payload.actionUrl || "#notifications";
  const options = {
    body: payload.body || "",
    icon: "assets/cyrra-logo.png",
    badge: "assets/cyrra-logo.png",
    tag: actionUrl,
    data: { actionUrl },
  };
  event.waitUntil(self.registration.showNotification(title, options));
});

self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  const actionUrl = event.notification.data?.actionUrl || "#notifications";
  event.waitUntil(
    self.clients.matchAll({ type: "window", includeUncontrolled: true }).then((clientList) => {
      for (const client of clientList) {
        if ("focus" in client) {
          client.postMessage({ type: "cyrra-push-navigate", actionUrl });
          return client.focus();
        }
      }
      if (self.clients.openWindow) {
        return self.clients.openWindow(actionUrl.startsWith("#") ? `/${actionUrl}` : `/#${actionUrl}`);
      }
    })
  );
});
