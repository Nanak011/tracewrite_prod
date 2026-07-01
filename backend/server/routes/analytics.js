const express = require("express");
const { requireAuth } = require("../controllers/authMiddleware");
const { requireProjectMember } = require("../controllers/projectAccess");
const analyticsController = require("../controllers/analyticsController");

const router = express.Router();

router.get("/:projectId/summary", requireAuth, requireProjectMember, analyticsController.projectSummary);
router.post("/:projectId/activity", requireAuth, analyticsController.trackActivity);
module.exports = router;



