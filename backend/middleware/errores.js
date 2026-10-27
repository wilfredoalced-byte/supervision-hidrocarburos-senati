function errorHandler(err, req, res, next) {
  if (res.headersSent) return next(err);
  console.error(err);
  res.status(Number.isInteger(err.status) && err.status >= 400 && err.status <= 599 ? err.status : 500)
    .json({ error: err.status && err.status < 500 ? err.message : 'Error interno del servidor' });
}
module.exports = { errorHandler };
