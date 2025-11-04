function formatMessage(level, message, data) {
    const timestamp = new Date().toISOString();
    const dataStr = data ? ` ${JSON.stringify(data, null, 2)}` : "";
    return `[${timestamp}] [${level}] ${message}${dataStr}`;
}
export const logger = {
    info(message, data) {
        const logMessage = formatMessage("INFO", message, data);
        console.log(logMessage);
    },
    error(message, error) {
        const logMessage = formatMessage("ERROR", message, error);
        console.error(logMessage);
    },
    // debug(message: string, data?: unknown) {
    //   const logMessage = formatMessage(
    //     "DEBUG",
    //     message,
    //     data,
    //   );
    //   console.log(logMessage);
    // },
};
