const request = require('supertest');
const mongoose = require('mongoose');
const app = require('../server');

describe('Task Management API Test Suite (Practical 12 - CI/CD Automated Tests)', () => {
  let authToken = '';
  const testEmail = `ci_test_${Date.now()}@charusat.edu.in`;
  const testPassword = 'Password@123';

  beforeAll(async () => {
    if (mongoose.connection.readyState === 0) {
      const MONGO_URI = process.env.MONGO_URI || 'mongodb://127.0.0.1:27017/task_manager_db';
      await mongoose.connect(MONGO_URI);
    } else if (mongoose.connection.readyState === 2) {
      await mongoose.connection.asPromise();
    }
  });

  afterAll(async () => {
    await mongoose.connection.close();
  });

  // Test 1: Root Health Check
  test('GET / should return 200 and API metadata', async () => {
    const res = await request(app).get('/');
    expect(res.statusCode).toBe(200);
    expect(res.body.success).toBe(true);
    expect(res.body.student).toBe('Priyansh Pravinbhai Panchal');
  });

  // Test 2: In-Memory Cache Stats (Practical 9)
  test('GET /cache/stats should return 200 and cache diagnostics', async () => {
    const res = await request(app).get('/cache/stats');
    expect(res.statusCode).toBe(200);
    expect(res.body.success).toBe(true);
    expect(res.body.data).toHaveProperty('stdTTL', 60);
  });

  // Test 3: Background Events Log (Practical 10)
  test('GET /events/log should return 200 and background audit log', async () => {
    const res = await request(app).get('/events/log');
    expect(res.statusCode).toBe(200);
    expect(res.body.success).toBe(true);
    expect(Array.isArray(res.body.events)).toBe(true);
  });

  // Test 4: Auth Guard (Practical 7)
  test('GET /tasks without token should be rejected with 401 Unauthorized', async () => {
    const res = await request(app).get('/tasks');
    expect(res.statusCode).toBe(401);
    expect(res.body.success).toBe(false);
  });

  // Test 5: User Registration & Login (Practical 7)
  test('POST /register and POST /login should authenticate and return JWT', async () => {
    // Register
    const regRes = await request(app)
      .post('/register')
      .send({ email: testEmail, password: testPassword });
    expect(regRes.statusCode).toBe(201);
    expect(regRes.body.success).toBe(true);

    // Login
    const loginRes = await request(app)
      .post('/login')
      .send({ email: testEmail, password: testPassword });
    expect(loginRes.statusCode).toBe(200);
    expect(loginRes.body.success).toBe(true);
    expect(loginRes.body).toHaveProperty('token');
    authToken = loginRes.body.token;
  });

  // Test 6: Authenticated GET /tasks returns 200 (Step 2)
  test('GET /tasks with valid Bearer token returns 200', async () => {
    const res = await request(app)
      .get('/tasks')
      .set('Authorization', `Bearer ${authToken}`);
    expect(res.statusCode).toBe(200);
    expect(res.body.success).toBe(true);
    expect(Array.isArray(res.body.data)).toBe(true);
  });

  // Test 7: Input Validation Middleware rejects empty title (Practical 7)
  test('POST /tasks with missing title should return 400 Bad Request', async () => {
    const res = await request(app)
      .post('/tasks')
      .set('Authorization', `Bearer ${authToken}`)
      .send({ description: 'Testing validation in CI', priority: 'high' });
    expect(res.statusCode).toBe(400);
    expect(res.body.error).toBe('Validation Error');
  });

  // Test 8: AI Task Description Generation with Fallback (Practical 13)
  test('POST /api/ai/generate-description should return 200 and generated suggestion', async () => {
    const res = await request(app)
      .post('/api/ai/generate-description')
      .send({ title: 'Complete Lab 13 AI Integration' });
    expect(res.statusCode).toBe(200);
    expect(res.body.success).toBe(true);
    expect(res.body).toHaveProperty('description');
    expect(res.body).toHaveProperty('priority');
    expect(res.body).toHaveProperty('disclaimer');
  });

  // Test 9: AI Task Description rejects empty title
  test('POST /api/ai/generate-description with empty title returns 400', async () => {
    const res = await request(app)
      .post('/api/ai/generate-description')
      .send({ title: '' });
    expect(res.statusCode).toBe(400);
    expect(res.body.error).toBe('Validation Error');
  });

  // Test 10: AI Rate Limiting (Supplementary Problem 2)
  test('POST /api/ai/generate-description enforces rate limiting on second immediate call', async () => {
    const res = await request(app)
      .post('/api/ai/generate-description')
      .send({ title: 'Another Rapid Task' });
    expect(res.statusCode).toBe(429);
    expect(res.body.error).toBe('Too Many Requests');
  });
});
