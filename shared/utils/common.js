/**
 * Shared utilities for Microsoft integrations
 * Common functions used across different Microsoft service integrations
 */

/**
 * Logger utility with consistent formatting across integrations
 */
class Logger {
    static log(message, level = 'INFO', context = '') {
        const timestamp = new Date().toISOString();
        const contextStr = context ? ` [${context}]` : '';
        console.log(`[${timestamp}] ${level}${contextStr}: ${message}`);
    }

    static info(message, context = '') {
        this.log(message, 'INFO', context);
    }

    static warn(message, context = '') {
        this.log(message, 'WARN', context);
    }

    static error(message, context = '') {
        this.log(message, 'ERROR', context);
    }

    static debug(message, context = '') {
        if (process.env.NODE_ENV === 'development') {
            this.log(message, 'DEBUG', context);
        }
    }
}

/**
 * Configuration management for different environments
 */
class Config {
    static getEnvironment() {
        return process.env.NODE_ENV || 'development';
    }

    static isDevelopment() {
        return this.getEnvironment() === 'development';
    }

    static isProduction() {
        return this.getEnvironment() === 'production';
    }

    static getServerUrl(defaultPort = 3000) {
        if (this.isDevelopment()) {
            return `http://localhost:${process.env.PORT || defaultPort}`;
        }
        return process.env.SERVER_URL || `https://localhost:${defaultPort}`;
    }
}

/**
 * Microsoft Graph API helper utilities
 */
class GraphApiHelper {
    static getAuthHeaders(accessToken) {
        return {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        };
    }

    static async makeRequest(url, options = {}) {
        try {
            const response = await fetch(url, {
                headers: {
                    'Content-Type': 'application/json',
                    ...options.headers
                },
                ...options
            });

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }

            return await response.json();
        } catch (error) {
            Logger.error(`Graph API request failed: ${error.message}`, 'GraphApi');
            throw error;
        }
    }
}

/**
 * Office.js common utilities
 */
class OfficeHelper {
    static async initializeOffice() {
        return new Promise((resolve, reject) => {
            if (typeof Office !== 'undefined') {
                Office.onReady((info) => {
                    Logger.info(`Office initialized: ${info.host} v${info.platform}`, 'Office');
                    resolve(info);
                });
            } else {
                reject(new Error('Office.js not available'));
            }
        });
    }

    static async runWithErrorHandling(operation, context = 'Office operation') {
        try {
            return await operation();
        } catch (error) {
            Logger.error(`${context} failed: ${error.message}`, 'Office');
            throw error;
        }
    }

    static formatError(error) {
        if (error.traceMessages) {
            return `${error.message} (${error.traceMessages.join(', ')})`;
        }
        return error.message || 'Unknown error';
    }
}

module.exports = {
    Logger,
    Config,
    GraphApiHelper,
    OfficeHelper
};
