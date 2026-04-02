const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode');
const { initializeApp, cert } = require('firebase-admin/app');
const { getFirestore, FieldValue } = require('firebase-admin/firestore');

// 1. Configuração do Firebase
const serviceAccount = require('./serviceAccountKey.json');

// Inicializa o Firebase de forma moderna
const app = initializeApp({
  credential: cert(serviceAccount)
});

// Conecta ao seu banco de dados específico
const db = getFirestore('ai-studio-ab1b8225-4f4e-4041-ae38-fe65699bfce4');

const client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: {
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    }
});

async function updateStatus(state, extra = {}) {
    try {
        await db.collection('status').doc('whatsapp').set({
            state,
            updatedAt: FieldValue.serverTimestamp(),
            ...extra
        }, { merge: true });
    } catch (e) { console.error("Erro ao atualizar status:", e); }
}

async function logToFirebase(message) {
    console.log(`WA_LOG: ${message}`);
    try {
        await db.collection('logs').add({
            message: `[WhatsApp] ${message}`,
            timestamp: FieldValue.serverTimestamp()
        });
    } catch (e) { console.error("Erro ao gravar log:", e); }
}

client.on('qr', (qr) => {
    console.log('QR CODE GERADO! Olhe no seu Dashboard Web.');
    qrcode.toDataURL(qr, (err, url) => {
        updateStatus('connecting', { qr: url });
    });
});

client.on('ready', () => {
    updateStatus('connected', { qr: null });
    logToFirebase('WhatsApp Conectado com sucesso!');
});

client.on('disconnected', (reason) => {
    updateStatus('disconnected', { qr: null });
    logToFirebase(`WhatsApp Desconectado: ${reason}`);
});

// Escuta mensagens para enviar
db.collection('messages').where('status', '==', 'pending').onSnapshot(snapshot => {
    snapshot.docChanges().forEach(async (change) => {
        if (change.type === 'added') {
            const data = change.doc.data();
            try {
                const chatId = data.to.includes('@c.us') ? data.to : `${data.to}@c.us`;
                await client.sendMessage(chatId, data.body);
                await change.doc.ref.update({ 
                    status: 'sent', 
                    sentAt: FieldValue.serverTimestamp() 
                });
            } catch (error) {
                console.error('Erro ao enviar mensagem:', error);
            }
        }
    });
});

console.log('>>> Iniciando WhatsApp Worker... Aguarde o QR Code no Dashboard.');
client.initialize();