const jwt = require('jsonwebtoken');

/**
 * Authentication Middleware (Practical 7 - Step 5 & Problem Definition)
 * Verifies JWT token from 'Authorization: Bearer <token>' header.
 * Attaches decoded payload (user id & email) to req.user.
 */
const authMiddleware = (req, res, next) => {
  const authHeader = req.headers.authorization;

  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return res.status(401).json({
      success: false,
      error: 'Unauthorized',
      message: 'Access denied. No token provided. Authorization header must be: Bearer <token>',
    });
  }

  const token = authHeader.split(' ')[1];

  if (!token) {
    return res.status(401).json({
      success: false,
      error: 'Unauthorized',
      message: 'Access denied. Token missing from Authorization header.',
    });
  }

  try {
    const secret = process.env.JWT_SECRET || 'charusat_awdf_jwt_secret_key_sem5_2026';
    const decoded = jwt.verify(token, secret);
    req.user = decoded;
    next();
  } catch (err) {
    if (err.name === 'TokenExpiredError') {
      return res.status(401).json({
        success: false,
        error: 'Token Expired',
        message: 'Your session has expired. Please log in again to continue.',
      });
    }
    return res.status(401).json({
      success: false,
      error: 'Invalid Token',
      message: 'Authentication failed. Provided JWT token is invalid or corrupt.',
    });
  }
};

module.exports = authMiddleware;
