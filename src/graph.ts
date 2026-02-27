import { Client } from '@microsoft/microsoft-graph-client';
import { getAccessToken } from './auth.js';
import 'isomorphic-fetch';

export const getGraphClient = (): Client => {
    return Client.init({
        authProvider: async (done) => {
            try {
                const token = await getAccessToken();
                done(null, token);
            } catch (error: any) {
                done(error, null);
            }
        }
    });
};