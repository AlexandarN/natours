const request = require('supertest');
const should = require('chai').should();
const { ObjectId } = require('mongoose').Types;
const app = require('../../app');
const { issueNewToken } = require('../../lib/jwtHandler');
const { addUser } = require('../helpers/userHelper');
const { addCompany } = require('../helpers/companyHelper');
const { addPOS } = require('../helpers/posHelper');
const { createGoalsForCompany } = require('../helpers/kpi2Helper');
const { Activity, KPI } = require('../../models');

describe.only('Edit Daily Goal', () => {
  let createdCompany;
  let createdPOS1;
  let createdPOS2;
  let createdAdmin;
  let createdInactiveAdmin;
  let createdManager;
  let createdUser;
  let createdYearQuery;
  let createdMonthQuery;
  let createdDayQuery;
  let createdYear;
  let createdMonth;
  let createdDay;

  before(async () => {
    createdCompany = await addCompany();
    [createdPOS1, createdPOS2] = await Promise.all([
      addPOS({ company: createdCompany._id }),
      addPOS({ company: createdCompany._id }),
    ]);
    [createdAdmin, createdInactiveAdmin, createdManager, createdUser] = await Promise.all([
      addUser({ role: 'Admin', company: createdCompany._id }),
      addUser({ role: 'Admin', company: createdCompany._id, isActive: false }),
      addUser({ role: 'Manager', company: createdCompany._id, pos: createdPOS1._id }),
      addUser({ role: 'User', company: createdCompany._id, pos: createdPOS1._id }),
      createGoalsForCompany({ company: createdCompany._id, posIDs: [createdPOS1._id, createdPOS2._id] }),
    ]);
    createdYearQuery = KPI.findOne(
      { date: new Date().getFullYear(), 'pointsOfSale.pos': createdPOS1._id },
      {
        _id: 1,
        status: 1,
        timeUnit: 1,
        date: 1,
        dayOfWeek: 1,
        totalMoney: 1,
        totalItems: 1,
        totalChecks: 1,
        totalVisitors: 1,
        'pointsOfSale.$': 1,
      },
    ).lean();
    createdMonthQuery = KPI.findOne(
      { date: `${new Date().getFullYear()}12`, 'pointsOfSale.pos': createdPOS1._id },
      {
        _id: 1,
        status: 1,
        timeUnit: 1,
        date: 1,
        dayOfWeek: 1,
        totalMoney: 1,
        totalItems: 1,
        totalChecks: 1,
        totalVisitors: 1,
        'pointsOfSale.$': 1,
      },
    ).lean();
    createdDayQuery = KPI.findOne(
      { date: `${new Date().getFullYear()}1231`, 'pointsOfSale.pos': createdPOS1._id },
      {
        _id: 1,
        status: 1,
        timeUnit: 1,
        date: 1,
        dayOfWeek: 1,
        totalMoney: 1,
        totalItems: 1,
        totalChecks: 1,
        totalVisitors: 1,
        'pointsOfSale.$': 1,
      },
    ).lean();
    [createdYear, createdMonth, createdDay] = await Promise.all([createdYearQuery, createdMonthQuery, createdDayQuery]);
  });

  it('POST /web/company/:companyId/pos/:posId/kpi2/goal/day/:dayId Should return unauthorized if token is not valid', (done) => {
    const invalidToken = issueNewToken({ _id: ObjectId() });
    const body = {};
    request(app)
      .post(`/api/v1/web/company/${createdCompany._id}/pos/${createdPOS1._id}/kpi2/goal/day/${createdDay.date}`)
      .set('Accept', 'application/json')
      .set('Authorization', `Bearer ${invalidToken}`)
      .send(body)
      .expect(401)
      .then((res) => {
        res.body.errorCode.should.equal(12);
        res.body.message.should.equal('Invalid credentials');
        done();
      })
      .catch(done);
  });

  it('POST /web/company/:companyId/pos/:posId/kpi2/goal/day/:dayId Should return unauthorized if user is not active', (done) => {
    const body = {};
    request(app)
      .post(`/api/v1/web/company/${createdCompany._id}/pos/${createdPOS1._id}/kpi2/goal/day/${createdDay.date}`)
      .set('Accept', 'application/json')
      .set('Authorization', `Bearer ${createdInactiveAdmin.token}`)
      .send(body)
      .expect(401)
      .then((res) => {
        res.body.errorCode.should.equal(12);
        res.body.message.should.equal('Invalid credentials');
        done();
      })
      .catch(done);
  });

  it('POST /web/company/:companyId/pos/:posId/kpi2/goal/day/:dayId Should return unauthorized if user role is not Admin or Manager', (done) => {
    const body = {};
    request(app)
      .post(`/api/v1/web/company/${createdCompany._id}/pos/${createdPOS1._id}/kpi2/goal/day/${createdDay.date}`)
      .set('Accept', 'application/json')
      .set('Authorization', `Bearer ${createdUser.token}`)
      .send(body)
      .expect(401)
      .then((res) => {
        res.body.errorCode.should.equal(12);
        res.body.message.should.equal('Invalid credentials');
        done();
      })
      .catch(done);
  });

  it('POST /web/company/:companyId/pos/:posId/kpi2/goal/day/:dayId Should return unauthorized if Manager tries to recalculate daily goals for another POS', (done) => {
    const body = {};
    request(app)
      .post(`/api/v1/web/company/${createdCompany._id}/pos/${createdPOS2._id}/kpi2/goal/day/${createdDay.date}`)
      .set('Accept', 'application/json')
      .set('Authorization', `Bearer ${createdManager.token}`)
      .send(body)
      .expect(401)
      .then((res) => {
        res.body.errorCode.should.equal(12);
        res.body.message.should.equal('Invalid credentials');
        done();
      })
      .catch(done);
  });

  it('POST /web/company/:companyId/pos/:posId/kpi2/goal/day/:dayId Should return missing parameters', (done) => {
    const body = {};
    request(app)
      .post(`/api/v1/web/company/${createdCompany._id}/pos/${createdPOS1._id}/kpi2/goal/day/${createdDay.date}`)
      .set('Accept', 'application/json')
      .set('Authorization', `Bearer ${createdManager.token}`)
      .send(body)
      .expect(400)
      .then((res) => {
        console.log(createdDay.pointsOfSale[0])

        res.body.errorCode.should.equal(2);
        res.body.message.should.equal('Missing parameters');
        done();
      })
      .catch(done);
  });

  it('POST /web/company/:companyId/pos/:posId/kpi2/goal/day/:dayId Should return not found if daily goal does not exist in database', (done) => {
    const body = {
      totalMoney: createdDay.pointsOfSale[0].totalMoney.goal + 100000,
      totalItems: createdDay.pointsOfSale[0].totalItems.goal + 200,
      totalChecks: createdDay.pointsOfSale[0].totalChecks.goal + 100,
      totalVisitors: createdDay.pointsOfSale[0].totalVisitors.goal - 300,
      averageItemPrice: createdDay.pointsOfSale[0].averageItemPrice.goal + 10,
      averageMoneyPerCheck: createdDay.pointsOfSale[0].averageMoneyPerCheck.goal + 100000,
      averageItemsPerCheck: createdDay.pointsOfSale[0].averageItemsPerCheck.goal + 100000,
      conversionRatio: createdDay.pointsOfSale[0].conversionRatio.goal + 100000,
    };
    request(app)
      .post(`/api/v1/web/company/${createdCompany._id}/pos/${createdPOS1._id}/kpi2/goal/day/${createdYear.date}`)
      .set('Accept', 'application/json')
      .set('Authorization', `Bearer ${createdAdmin.token}`)
      .send(body)
      .expect(404)
      .then((res) => {

        res.body.errorCode.should.equal(4);
        res.body.message.should.equal('Not Found');
        done();
      })
      .catch(done);
  });

  it('POST /web/company/:companyId/pos/:posId/kpi2/goal/day/:dayId Should return not found if day has passed', (done) => {
    const body = {
      newTotalMoneyGoal: createdDay.pointsOfSale[0].totalMoney.goal + 150000,
      totalMoneyCorrelation: 20,
      totalChecksCorrelation: 20,
      averageMoneyPerCheckCorrelation: 20,
      saveChanges: false,
    };
    request(app)
      .post(`/api/v1/web/company/${createdCompany._id}/pos/${createdPOS1._id}/kpi2/goal/day/${createdYear.date}0101`)
      .set('Accept', 'application/json')
      .set('Authorization', `Bearer ${createdAdmin.token}`)
      .send(body)
      .expect(404)
      .then((res) => {
        res.body.errorCode.should.equal(4);
        res.body.message.should.equal('Not Found');
        done();
      })
      .catch(done);
  });

  it('POST /web/company/:companyId/pos/:posId/kpi2/goal/day/:dayId Should successfully edit daily goals and update respective total money values in DB (Admin)', (done) => {
    const body = {
      newTotalMoneyGoal: createdDay.pointsOfSale[0].totalMoney.goal + 200000,
      totalMoneyCorrelation: 20,
      totalChecksCorrelation: 20,
      averageMoneyPerCheckCorrelation: 20,
      saveChanges: false,
    };
    request(app)
      .post(`/api/v1/web/company/${createdCompany._id}/pos/${createdPOS1._id}/kpi2/goal/day/${createdDay.date}`)
      .set('Accept', 'application/json')
      .set('Authorization', `Bearer ${createdAdmin.token}`)
      .send(body)
      .expect(200)
      .then((res) => {
        res.body.message.should.equal('Successfully recalculated daily goal');
        res.body.results.totalMoney.initialGoal.should.not.equal(res.body.results.totalMoney.goal);
        res.body.results.totalMoney.goal.should.equal(body.newTotalMoneyGoal);
        res.body.results.totalMoney.goal.should.not.equal(res.body.results.totalMoney.oldGoal);
        res.body.results.totalItems.initialGoal.should.not.equal(res.body.results.totalItems.goal);
        res.body.results.totalItems.goal.should.not.equal(res.body.results.totalItems.oldGoal);
        res.body.results.totalChecks.initialGoal.should.not.equal(res.body.results.totalChecks.goal);
        res.body.results.totalChecks.goal.should.not.equal(res.body.results.totalChecks.oldGoal);
        res.body.results.totalVisitors.initialGoal.should.not.equal(res.body.results.totalVisitors.goal);
        res.body.results.totalVisitors.goal.should.not.equal(res.body.results.totalVisitors.oldGoal);
        res.body.results.averageItemPrice.initialGoal.should.not.equal(res.body.results.averageItemPrice.goal);
        res.body.results.averageItemPrice.goal.should.not.equal(res.body.results.averageItemPrice.oldGoal);
        res.body.results.averageMoneyPerCheck.initialGoal.should.not.equal(res.body.results.averageMoneyPerCheck.goal);
        res.body.results.averageMoneyPerCheck.goal.should.not.equal(res.body.results.averageMoneyPerCheck.oldGoal);
        res.body.results.averageItemsPerCheck.initialGoal.should.not.equal(res.body.results.averageItemsPerCheck.goal);
        res.body.results.averageItemsPerCheck.goal.should.not.equal(res.body.results.averageItemsPerCheck.oldGoal);
        res.body.results.conversionRatio.initialGoal.should.not.equal(res.body.results.conversionRatio.goal);
        res.body.results.conversionRatio.goal.should.not.equal(res.body.results.conversionRatio.oldGoal);
        return Promise.all([
          createdYearQuery,
          createdMonthQuery,
          createdDayQuery,
          Activity.findOne({ user: createdAdmin.user._id, 'changes.date': createdDay.date }),
        ])
          .then(([dbYear, dbMonth, dbDay, dbActivity]) => {
            dbYear.totalMoney.initialGoal.should.equal(createdYear.totalMoney.initialGoal);
            dbYear.totalMoney.goal.should.equal(createdYear.totalMoney.goal);
            dbYear.totalItems.initialGoal.should.equal(createdYear.totalItems.initialGoal);
            dbYear.totalItems.goal.should.equal(createdYear.totalItems.goal);
            dbYear.totalChecks.initialGoal.should.equal(createdYear.totalChecks.initialGoal);
            dbYear.totalChecks.goal.should.equal(createdYear.totalChecks.goal);
            dbYear.totalVisitors.initialGoal.should.equal(createdYear.totalVisitors.initialGoal);
            dbYear.totalVisitors.goal.should.equal(createdYear.totalVisitors.goal);
            dbYear.pointsOfSale[0].totalMoney.initialGoal.should.equal(createdYear.pointsOfSale[0].totalMoney.initialGoal);
            dbYear.pointsOfSale[0].totalMoney.goal.should.equal(createdYear.pointsOfSale[0].totalMoney.goal);
            dbYear.pointsOfSale[0].totalItems.initialGoal.should.equal(createdYear.pointsOfSale[0].totalItems.initialGoal);
            dbYear.pointsOfSale[0].totalItems.goal.should.equal(createdYear.pointsOfSale[0].totalItems.goal);
            dbYear.pointsOfSale[0].totalChecks.initialGoal.should.equal(createdYear.pointsOfSale[0].totalChecks.initialGoal);
            dbYear.pointsOfSale[0].totalChecks.goal.should.equal(createdYear.pointsOfSale[0].totalChecks.goal);
            dbYear.pointsOfSale[0].totalVisitors.initialGoal.should.equal(createdYear.pointsOfSale[0].totalVisitors.initialGoal);
            dbYear.pointsOfSale[0].totalVisitors.goal.should.equal(createdYear.pointsOfSale[0].totalVisitors.goal);
            dbMonth.totalMoney.initialGoal.should.equal(createdMonth.totalMoney.initialGoal);
            dbMonth.totalMoney.goal.should.equal(createdMonth.totalMoney.goal);
            dbMonth.totalItems.initialGoal.should.equal(createdMonth.totalItems.initialGoal);
            dbMonth.totalItems.goal.should.equal(createdMonth.totalItems.goal);
            dbMonth.totalChecks.initialGoal.should.equal(createdMonth.totalChecks.initialGoal);
            dbMonth.totalChecks.goal.should.equal(createdMonth.totalChecks.goal);
            dbMonth.totalVisitors.initialGoal.should.equal(createdMonth.totalVisitors.initialGoal);
            dbMonth.totalVisitors.goal.should.equal(createdMonth.totalVisitors.goal);
            dbMonth.pointsOfSale[0].totalMoney.initialGoal.should.equal(createdMonth.pointsOfSale[0].totalMoney.initialGoal);
            dbMonth.pointsOfSale[0].totalMoney.goal.should.equal(createdMonth.pointsOfSale[0].totalMoney.goal);
            dbMonth.pointsOfSale[0].totalItems.initialGoal.should.equal(createdMonth.pointsOfSale[0].totalItems.initialGoal);
            dbMonth.pointsOfSale[0].totalItems.goal.should.equal(createdMonth.pointsOfSale[0].totalItems.goal);
            dbMonth.pointsOfSale[0].totalChecks.initialGoal.should.equal(createdMonth.pointsOfSale[0].totalChecks.initialGoal);
            dbMonth.pointsOfSale[0].totalChecks.goal.should.equal(createdMonth.pointsOfSale[0].totalChecks.goal);
            dbMonth.pointsOfSale[0].totalVisitors.initialGoal.should.equal(createdMonth.pointsOfSale[0].totalVisitors.initialGoal);
            dbMonth.pointsOfSale[0].totalVisitors.goal.should.equal(createdMonth.pointsOfSale[0].totalVisitors.goal);
            dbDay.totalMoney.initialGoal.should.equal(createdDay.totalMoney.initialGoal);
            dbDay.totalMoney.goal.should.equal(createdDay.totalMoney.goal);
            dbDay.totalItems.initialGoal.should.equal(createdDay.totalItems.initialGoal);
            dbDay.totalItems.goal.should.equal(createdDay.totalItems.goal);
            dbDay.totalChecks.initialGoal.should.equal(createdDay.totalChecks.initialGoal);
            dbDay.totalChecks.goal.should.equal(createdDay.totalChecks.goal);
            dbDay.totalVisitors.initialGoal.should.equal(createdDay.totalVisitors.initialGoal);
            dbDay.totalVisitors.goal.should.equal(createdDay.totalVisitors.goal);
            dbDay.pointsOfSale[0].totalMoney.initialGoal.should.equal(createdDay.pointsOfSale[0].totalMoney.initialGoal);
            dbDay.pointsOfSale[0].totalMoney.goal.should.equal(res.body.results.totalMoney.oldGoal);
            dbDay.pointsOfSale[0].totalItems.initialGoal.should.equal(createdDay.pointsOfSale[0].totalItems.initialGoal);
            dbDay.pointsOfSale[0].totalItems.goal.should.equal(res.body.results.totalItems.oldGoal);
            dbDay.pointsOfSale[0].totalChecks.initialGoal.should.equal(createdDay.pointsOfSale[0].totalChecks.initialGoal);
            dbDay.pointsOfSale[0].totalChecks.goal.should.equal(res.body.results.totalChecks.oldGoal);
            dbDay.pointsOfSale[0].totalVisitors.initialGoal.should.equal(createdDay.pointsOfSale[0].totalVisitors.initialGoal);
            dbDay.pointsOfSale[0].totalVisitors.goal.should.equal(res.body.results.totalVisitors.oldGoal);
            should.not.exist(dbActivity);
            done();
          });
      })
      .catch(done);
  });

  it('POST /web/company/:companyId/pos/:posId/kpi2/goal/day/:dayId Should successfully edit daily goals and update respective total money values in DB (Manager)', (done) => {
    const body = {
      newTotalMoneyGoal: createdDay.pointsOfSale[0].totalMoney.goal + 250000,
      totalMoneyCorrelation: 20,
      totalChecksCorrelation: 20,
      averageMoneyPerCheckCorrelation: 20,
      saveChanges: true,
    };
    request(app)
      .post(`/api/v1/web/company/${createdCompany._id}/pos/${createdPOS1._id}/kpi2/goal/day/${createdDay.date}`)
      .set('Accept', 'application/json')
      .set('Authorization', `Bearer ${createdManager.token}`)
      .send(body)
      .expect(200)
      .then((res) => {
        res.body.message.should.equal('Successfully recalculated daily goal');
        res.body.results.totalMoney.initialGoal.should.not.equal(res.body.results.totalMoney.goal);
        res.body.results.totalMoney.goal.should.equal(body.newTotalMoneyGoal);
        res.body.results.totalMoney.goal.should.not.equal(res.body.results.totalMoney.oldGoal);
        res.body.results.totalItems.initialGoal.should.not.equal(res.body.results.totalItems.goal);
        res.body.results.totalItems.goal.should.not.equal(res.body.results.totalItems.oldGoal);
        res.body.results.totalChecks.initialGoal.should.not.equal(res.body.results.totalChecks.goal);
        res.body.results.totalChecks.goal.should.not.equal(res.body.results.totalChecks.oldGoal);
        res.body.results.totalVisitors.initialGoal.should.not.equal(res.body.results.totalVisitors.goal);
        res.body.results.totalVisitors.goal.should.not.equal(res.body.results.totalVisitors.oldGoal);
        res.body.results.averageItemPrice.initialGoal.should.not.equal(res.body.results.averageItemPrice.goal);
        res.body.results.averageItemPrice.goal.should.not.equal(res.body.results.averageItemPrice.oldGoal);
        res.body.results.averageMoneyPerCheck.initialGoal.should.not.equal(res.body.results.averageMoneyPerCheck.goal);
        res.body.results.averageMoneyPerCheck.goal.should.not.equal(res.body.results.averageMoneyPerCheck.oldGoal);
        res.body.results.averageItemsPerCheck.initialGoal.should.not.equal(res.body.results.averageItemsPerCheck.goal);
        res.body.results.averageItemsPerCheck.goal.should.not.equal(res.body.results.averageItemsPerCheck.oldGoal);
        res.body.results.conversionRatio.initialGoal.should.not.equal(res.body.results.conversionRatio.goal);
        res.body.results.conversionRatio.goal.should.not.equal(res.body.results.conversionRatio.oldGoal);
        const totalMoneyDiff = res.body.results.totalMoney.goal - res.body.results.totalMoney.oldGoal;
        const totalItemsDiff = res.body.results.totalItems.goal - res.body.results.totalItems.oldGoal;
        const totalChecksDiff = res.body.results.totalChecks.goal - res.body.results.totalChecks.oldGoal;
        const totalVisitorsDiff = res.body.results.totalVisitors.goal - res.body.results.totalVisitors.oldGoal;
        return Promise.all([
          createdYearQuery,
          createdMonthQuery,
          createdDayQuery,
          Activity.findOne({ user: createdManager.user._id, 'changes.date': createdDay.date }),
        ])
          .then(([dbYear, dbMonth, dbDay, dbActivity]) => {
            dbYear.totalMoney.initialGoal.should.equal(createdYear.totalMoney.initialGoal);
            dbYear.totalMoney.goal.should.equal(createdYear.totalMoney.goal);
            dbYear.totalItems.initialGoal.should.equal(createdYear.totalItems.initialGoal);
            dbYear.totalItems.goal.should.equal(createdYear.totalItems.goal);
            dbYear.totalChecks.initialGoal.should.equal(createdYear.totalChecks.initialGoal);
            dbYear.totalChecks.goal.should.equal(createdYear.totalChecks.goal);
            dbYear.totalVisitors.initialGoal.should.equal(createdYear.totalVisitors.initialGoal);
            dbYear.totalVisitors.goal.should.equal(createdYear.totalVisitors.goal);
            dbYear.pointsOfSale[0].totalMoney.initialGoal.should.equal(createdYear.pointsOfSale[0].totalMoney.initialGoal);
            dbYear.pointsOfSale[0].totalMoney.goal.should.equal(createdYear.pointsOfSale[0].totalMoney.goal);
            dbYear.pointsOfSale[0].totalItems.initialGoal.should.equal(createdYear.pointsOfSale[0].totalItems.initialGoal);
            dbYear.pointsOfSale[0].totalItems.goal.should.equal(createdYear.pointsOfSale[0].totalItems.goal);
            dbYear.pointsOfSale[0].totalChecks.initialGoal.should.equal(createdYear.pointsOfSale[0].totalChecks.initialGoal);
            dbYear.pointsOfSale[0].totalChecks.goal.should.equal(createdYear.pointsOfSale[0].totalChecks.goal);
            dbYear.pointsOfSale[0].totalVisitors.initialGoal.should.equal(createdYear.pointsOfSale[0].totalVisitors.initialGoal);
            dbYear.pointsOfSale[0].totalVisitors.goal.should.equal(createdYear.pointsOfSale[0].totalVisitors.goal);
            dbMonth.totalMoney.initialGoal.should.equal(createdMonth.totalMoney.initialGoal);
            dbMonth.totalMoney.goal.should.equal(createdMonth.totalMoney.goal);
            dbMonth.totalItems.initialGoal.should.equal(createdMonth.totalItems.initialGoal);
            dbMonth.totalItems.goal.should.equal(createdMonth.totalItems.goal);
            dbMonth.totalChecks.initialGoal.should.equal(createdMonth.totalChecks.initialGoal);
            dbMonth.totalChecks.goal.should.equal(createdMonth.totalChecks.goal);
            dbMonth.totalVisitors.initialGoal.should.equal(createdMonth.totalVisitors.initialGoal);
            dbMonth.totalVisitors.goal.should.equal(createdMonth.totalVisitors.goal);
            dbMonth.pointsOfSale[0].totalMoney.initialGoal.should.equal(createdMonth.pointsOfSale[0].totalMoney.initialGoal);
            dbMonth.pointsOfSale[0].totalMoney.goal.should.equal(createdMonth.pointsOfSale[0].totalMoney.goal);
            dbMonth.pointsOfSale[0].totalItems.initialGoal.should.equal(createdMonth.pointsOfSale[0].totalItems.initialGoal);
            dbMonth.pointsOfSale[0].totalItems.goal.should.equal(createdMonth.pointsOfSale[0].totalItems.goal);
            dbMonth.pointsOfSale[0].totalChecks.initialGoal.should.equal(createdMonth.pointsOfSale[0].totalChecks.initialGoal);
            dbMonth.pointsOfSale[0].totalChecks.goal.should.equal(createdMonth.pointsOfSale[0].totalChecks.goal);
            dbMonth.pointsOfSale[0].totalVisitors.initialGoal.should.equal(createdMonth.pointsOfSale[0].totalVisitors.initialGoal);
            dbMonth.pointsOfSale[0].totalVisitors.goal.should.equal(createdMonth.pointsOfSale[0].totalVisitors.goal);
            dbDay.totalMoney.initialGoal.should.equal(createdDay.totalMoney.initialGoal);
            dbDay.totalMoney.goal.should.equal(createdDay.totalMoney.goal + totalMoneyDiff);
            dbDay.totalItems.initialGoal.should.equal(createdDay.totalItems.initialGoal);
            dbDay.totalItems.goal.should.equal(createdDay.totalItems.goal + totalItemsDiff);
            dbDay.totalChecks.initialGoal.should.equal(createdDay.totalChecks.initialGoal);
            dbDay.totalChecks.goal.should.equal(createdDay.totalChecks.goal + totalChecksDiff);
            dbDay.totalVisitors.initialGoal.should.equal(createdDay.totalVisitors.initialGoal);
            dbDay.totalVisitors.goal.should.equal(createdDay.totalVisitors.goal + totalVisitorsDiff);
            dbDay.pointsOfSale[0].totalMoney.initialGoal.should.equal(createdDay.pointsOfSale[0].totalMoney.initialGoal);
            dbDay.pointsOfSale[0].totalMoney.goal.should.equal(createdDay.pointsOfSale[0].totalMoney.goal + totalMoneyDiff);
            dbDay.pointsOfSale[0].totalMoney.goal.should.equal(res.body.results.totalMoney.goal);
            dbDay.pointsOfSale[0].totalItems.initialGoal.should.equal(createdDay.pointsOfSale[0].totalItems.initialGoal);
            dbDay.pointsOfSale[0].totalItems.goal.should.equal(createdDay.pointsOfSale[0].totalItems.goal + totalItemsDiff);
            dbDay.pointsOfSale[0].totalItems.goal.should.equal(res.body.results.totalItems.goal);
            dbDay.pointsOfSale[0].totalChecks.initialGoal.should.equal(createdDay.pointsOfSale[0].totalChecks.initialGoal);
            dbDay.pointsOfSale[0].totalChecks.goal.should.equal(createdDay.pointsOfSale[0].totalChecks.goal + totalChecksDiff);
            dbDay.pointsOfSale[0].totalChecks.goal.should.equal(res.body.results.totalChecks.goal);
            dbDay.pointsOfSale[0].totalVisitors.initialGoal.should.equal(createdDay.pointsOfSale[0].totalVisitors.initialGoal);
            dbDay.pointsOfSale[0].totalVisitors.goal.should.equal(createdDay.pointsOfSale[0].totalVisitors.goal + totalVisitorsDiff);
            dbDay.pointsOfSale[0].totalVisitors.goal.should.equal(res.body.results.totalVisitors.goal);
            dbActivity.eventCode.should.equal(1);
            done();
          });
      })
      .catch(done);
  });
});
