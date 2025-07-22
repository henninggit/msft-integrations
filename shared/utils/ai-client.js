/**
 * AI Integration Helper for Office Add-ins
 * Provides easy access to AI services from Office applications
 */

class AIServiceClient {
    constructor(gatewayUrl = 'http://localhost:8000') {
        this.gatewayUrl = gatewayUrl;
        this.defaultModel = 'gpt-3.5-turbo';
    }

    /**
     * Send a chat completion request to the LLM Gateway
     */
    async chatCompletion(prompt, options = {}) {
        const requestData = {
            model: options.model || this.defaultModel,
            messages: [{ role: 'user', content: prompt }],
            temperature: options.temperature || 0.7,
            max_tokens: options.maxTokens || null,
            provider: options.provider || null
        };

        try {
            const response = await fetch(`${this.gatewayUrl}/v1/chat/completions`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(requestData)
            });

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }

            const data = await response.json();
            return {
                content: data.choices[0].message.content,
                model: data.model,
                usage: data.usage
            };
        } catch (error) {
            console.error('AI Service Error:', error);
            throw new Error(`Failed to get AI response: ${error.message}`);
        }
    }

    /**
     * Analyze document content
     */
    async analyzeDocument(content) {
        const prompt = `
Analyze the following document content and provide insights:

Content: ${content}

Please provide:
1. A brief summary
2. Key themes or topics
3. Suggested improvements
4. Tone and style assessment
        `.trim();

        return await this.chatCompletion(prompt);
    }

    /**
     * Improve text quality
     */
    async improveText(text) {
        const prompt = `
Please improve the following text while maintaining its original meaning:

Original text: ${text}

Focus on:
- Grammar and spelling
- Clarity and readability
- Professional tone
- Conciseness
        `.trim();

        return await this.chatCompletion(prompt);
    }

    /**
     * Generate content based on requirements
     */
    async generateContent(requirements) {
        const prompt = `
Generate content based on the following requirements:

${requirements}

Please create well-structured, professional content that meets these requirements.
        `.trim();

        return await this.chatCompletion(prompt);
    }

    /**
     * Analyze data and provide insights
     */
    async analyzeData(data, dataType = 'general') {
        const prompt = `
Analyze the following ${dataType} data and provide insights:

Data: ${data}

Please provide:
1. Key patterns or trends
2. Notable observations
3. Potential implications
4. Recommendations for action
        `.trim();

        return await this.chatCompletion(prompt);
    }

    /**
     * Check if AI services are available
     */
    async checkStatus() {
        try {
            const response = await fetch(`${this.gatewayUrl}/`);
            if (response.ok) {
                return await response.json();
            }
            return { status: 'unavailable' };
        } catch (error) {
            return { status: 'error', message: error.message };
        }
    }

    /**
     * Switch AI provider
     */
    async switchProvider(provider) {
        try {
            const response = await fetch(`${this.gatewayUrl}/v1/provider/switch`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ provider })
            });

            return response.ok;
        } catch (error) {
            console.error('Failed to switch provider:', error);
            return false;
        }
    }
}

// Export for use in Office add-ins
if (typeof module !== 'undefined' && module.exports) {
    module.exports = AIServiceClient;
}

// Make available globally for browser environments
if (typeof window !== 'undefined') {
    window.AIServiceClient = AIServiceClient;
}
