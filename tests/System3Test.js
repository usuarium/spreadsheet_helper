const sinon = require('sinon')
const assert = require('assert')
const path = require('path')
const NodeVMHelper = require('./NodeVMHelper.js')
let system3

before(async () => {
    let vmHelper = new NodeVMHelper()
    await vmHelper.loadContext(path.resolve(__dirname, '../src'), {
        console: console
    })
    system3 = vmHelper.instantiateClass('System3')
});

describe('tests System3', function() {

    it('tests System3.hasShelfmarkColumn', function() {
        let sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        system3.loadSheetHeaders()
        system3.determinateSheetState()
        assert.deepEqual(system3.hasShelfmarkColumn, false)
        sheetWrapperMock.verify()

        // worng position
        sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','SHELFMARK','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        system3.loadSheetHeaders()
        system3.determinateSheetState()
        assert.deepEqual(system3.hasShelfmarkColumn, false)
        sheetWrapperMock.verify()

        sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','SHELFMARK','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        system3.loadSheetHeaders()
        system3.determinateSheetState()
        assert.deepEqual(system3.hasShelfmarkColumn, true)
        sheetWrapperMock.verify()
    });


    it('tests System3.hasSourceColumn', function() {
        let sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        system3.loadSheetHeaders()
        system3.determinateSheetState()
        assert.deepEqual(system3.hasSourceColumn, false)
        sheetWrapperMock.verify()

        // worng position
        sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','SOURCE','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        system3.loadSheetHeaders()
        system3.determinateSheetState()
        assert.deepEqual(system3.hasSourceColumn, false)
        sheetWrapperMock.verify()

        sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','SHELFMARK','SOURCE','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        system3.loadSheetHeaders()
        system3.determinateSheetState()
        assert.deepEqual(system3.hasSourceColumn, true)
        sheetWrapperMock.verify()
    });


    it('tests System3.hasTopicsColumn', function() {
        let sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        system3.loadSheetHeaders()
        system3.determinateSheetState()
        assert.deepEqual(system3.hasTopicsColumn, false)
        sheetWrapperMock.verify()

        // worng position
        sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','TOPICS','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        system3.loadSheetHeaders()
        system3.determinateSheetState()
        assert.deepEqual(system3.hasTopicsColumn, false)
        sheetWrapperMock.verify()

        sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','TOPICS','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        system3.loadSheetHeaders()
        system3.determinateSheetState()
        assert.deepEqual(system3.hasTopicsColumn, true)
        sheetWrapperMock.verify()
    });


    it('tests System3.hasPageLinkColumn', function() {
        let sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        system3.loadSheetHeaders()
        system3.determinateSheetState()
        assert.deepEqual(system3.hasPageLinkColumn, false)
        sheetWrapperMock.verify()

        // worng position
        sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE LINK','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        system3.loadSheetHeaders()
        system3.determinateSheetState()
        assert.deepEqual(system3.hasPageLinkColumn, false)
        sheetWrapperMock.verify()

        sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','PAGE LINK','REMARK','MADE BY'
        ])

        system3.loadSheetHeaders()
        system3.determinateSheetState()
        assert.deepEqual(system3.hasPageLinkColumn, true)
        sheetWrapperMock.verify()
    });


    it('tests System3.getHeadersTemplate', function() {
        let sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        system3.loadSheetHeaders()
        system3.determinateSheetState()
        assert.deepEqual(system3.getHeadersTemplate(), [
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','SHELFMARK','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        system3.loadSheetHeaders()
        system3.determinateSheetState()
        assert.deepEqual(system3.getHeadersTemplate(), [
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])
        sheetWrapperMock.verify()


        sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','SHELFMARK','SOURCE','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','TOPICS','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','PAGE LINK','REMARK','MADE BY'
        ])

        system3.loadSheetHeaders()
        system3.determinateSheetState()
        assert.deepEqual(system3.getHeadersTemplate(), [
            'ID','SHELFMARK','SOURCE','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','TOPICS','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','PAGE LINK','REMARK','MADE BY'
        ])
        sheetWrapperMock.verify()
    });


    it('tests System3.insertHeaderAfter with shelfmark header', function() {
        let sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])
        sheetWrapperMock.expects('insertColumnAtPosition').once().withArgs(1)
        sheetWrapperMock.expects('setHeaderValueAtPosition').once().withArgs(2, 'SHELFMARK')

        system3.loadSheetHeaders()
        system3.determinateSheetState()

        system3.insertHeaderAfter('SHELFMARK', 'ID')

        assert.deepEqual(system3.headers, [
            'ID','SHELFMARK','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        sheetWrapperMock.verify()
    });


    it('tests System3.insertHeaderAfter with source header', function() {
        let sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])
        sheetWrapperMock.expects('insertColumnAtPosition').once().withArgs(1)
        sheetWrapperMock.expects('setHeaderValueAtPosition').once().withArgs(2, 'SHELFMARK')

        sheetWrapperMock.expects('insertColumnAtPosition').once().withArgs(2)
        sheetWrapperMock.expects('setHeaderValueAtPosition').once().withArgs(3, 'SOURCE')

        system3.loadSheetHeaders()
        system3.determinateSheetState()

        system3.insertHeaderAfter('SHELFMARK', 'ID')
        system3.insertHeaderAfter('SOURCE', 'SHELFMARK')

        assert.deepEqual(system3.headers, [
            'ID','SHELFMARK','SOURCE','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        sheetWrapperMock.verify()
    });


    it('tests System3.insertHeaderAfter with topics header', function() {
        let sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])
        sheetWrapperMock.expects('insertColumnAtPosition').once().withArgs(8)
        sheetWrapperMock.expects('setHeaderValueAtPosition').once().withArgs(9, 'TOPICS')

        system3.loadSheetHeaders()
        system3.determinateSheetState()

        system3.insertHeaderAfter('TOPICS', 'COMMUNE/VOTIVE')

        assert.deepEqual(system3.headers, [
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','TOPICS','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])

        sheetWrapperMock.verify()
    });


    it('tests System3.insertHeaderAfter with page link header', function() {
        let sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])
        sheetWrapperMock.expects('insertColumnAtPosition').once().withArgs(19)
        sheetWrapperMock.expects('setHeaderValueAtPosition').once().withArgs(20, 'PAGE LINK')

        system3.loadSheetHeaders()
        system3.determinateSheetState()

        system3.insertHeaderAfter('PAGE LINK', 'PAGE NUMBER (ORIGINAL)')

        assert.deepEqual(system3.headers, [
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','PAGE LINK','REMARK','MADE BY'
        ])

        sheetWrapperMock.verify()
    });


    it('tests System3.insertHeaderAfter with page link header', function() {
        let sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getHeaders').atLeast(1).returns([
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','REMARK','MADE BY'
        ])
        sheetWrapperMock.expects('insertColumnAtPosition').once().withArgs(19)
        sheetWrapperMock.expects('setHeaderValueAtPosition').once().withArgs(20, 'PAGE LINK')

        system3.loadSheetHeaders()
        system3.determinateSheetState()

        system3.insertHeaderAfter('PAGE LINK', 'PAGE NUMBER (ORIGINAL)')

        assert.deepEqual(system3.headers, [
            'ID','TYPE','PART','SEASON/MONTH','WEEK','DAY','FEAST','COMMUNE/VOTIVE','MASS/HOUR',
            'CEREMONY','MODULE','SEQUENCE','RUBRICS','LAYER','GENRE','SERIES','ITEM',
            'PAGE NUMBER (DIGITAL)','PAGE NUMBER (ORIGINAL)','PAGE LINK','REMARK','MADE BY'
        ])

        sheetWrapperMock.verify()
    });


    it('tests System3.getValidValues', function() {
        assert.deepEqual(system3.getValidValues('type'), ['MISS', 'OFF', 'RIT'])
        
        assert.deepEqual(system3.getValidValues('part', {type: 'MISS'}), ['T', 'S', 'C', 'V'])
        assert.deepEqual(system3.getValidValues('part', {type: 'OFF'}), ['T', 'S', 'C', 'V'])
        assert.deepEqual(system3.getValidValues('part', {type: 'RIT'}), ['O'])
        assert.deepEqual(system3.getValidValues('part', {type: 'invalid'}), null)

        assert.deepEqual(system3.getValidValues('season_month', {part: 'T'}), ['Adv', 'Nat', 'Ep', 'LXX', 'Qu', 'Pasc', 'Pent', 'Trin', 'Gen'])
        assert.deepEqual(system3.getValidValues('season_month', {part: 'S'}), ['Adv', 'Nat', 'Ep', 'LXX', 'Qu', 'Pasc', 'Pent', 'Trin', 'Gen'])
        assert.deepEqual(system3.getValidValues('season_month', {part: 'C'}), '')
        assert.deepEqual(system3.getValidValues('season_month', {part: 'V'}), '')

        assert.deepEqual(system3.getValidValues('week', {part: 'T'}), ['H0', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'H7', 'H8', 'H9', 'H10', 'H11', 'H12', 'H13', 'H14', 'H15', 'H16', 'H17', 'H18', 'H19', 'H20', 'H21', 'H22', 'H23', 'H24', 'H25', 'H26', 'H27', 'QuT', 'Gen'])
        assert.deepEqual(system3.getValidValues('week', {part: 'S'}), '')
        assert.deepEqual(system3.getValidValues('week', {part: 'C'}), '')
        assert.deepEqual(system3.getValidValues('week', {part: 'V'}), '')

        let regexp = system3.getValidValues('day', {part: 'T'})
        assert(regexp.constructor.name === 'RegExp')
        for (let i = 1; i <= 31; i++) {
            assert(regexp.exec(i))
        }

        assert(regexp.exec(0) === null)
        assert(regexp.exec(32) === null)
        assert(regexp.exec(100) === null)
        
        assert.deepEqual(system3.getValidValues('day', {part: 'S'}), ['D', 'ff', 'f2', 'f3', 'f4', 'f5', 'f6', 'S', 'Gen', 'F'])
        assert.deepEqual(system3.getValidValues('day', {part: 'C'}), '')
        assert.deepEqual(system3.getValidValues('day', {part: 'V'}), '')
        
        assert.deepEqual(system3.getValidValues('mass_hour', {type: 'MISS'}), ['', 'M1', 'M2', 'M3'])
        assert.deepEqual(system3.getValidValues('mass_hour', {type: 'OFF'}), ['V1', 'C1', 'M', 'N1', 'N2', 'N3', 'L', 'I', 'I+', 'III', 'VI', 'IX', 'V', 'V2', 'C', 'C2'])
        assert.deepEqual(system3.getValidValues('mass_hour', {type: 'RIT'}), '')

        assert.deepEqual(system3.getValidValues('ceremony', {type: 'MISS'}), ['', 'Mass Propers'])
        assert.deepEqual(system3.getValidValues('ceremony', {type: 'OFF'}), ['', 'Office Propers'])
        assert.deepEqual(system3.getValidValues('ceremony', {type: 'RIT'}), '')
        
        regexp = system3.getValidValues('sequence')
        assert(regexp.constructor.name === 'RegExp')
        assert(regexp.exec(1))
        assert(regexp.exec(2))
        assert(regexp.exec('') === null)

        assert.deepEqual(system3.getValidValues('layer'), ['S/C', 'G/A', 'L', 'R']);

        assert.deepEqual(system3.getValidValues('genre', {type: 'MISS', layer: 'G/A'}), [
            // romana
            'Intr', 'IntrV', 'IntrTrop', 'Tr', 'Gr', 'GrV', 'H', 'All', 'AllV', 'Seq', 'Off', 'OffV', 
            'OffTrop', 'Comm', 'CommV', 'CommTrop', 
            // ambrosiana
            'Ingressa ', 'Psalmellus', 'PsalmellusV', 'Cantus', 'CantusV', 'V. in Alleluia', 
            'Post evangelium', 'Offerenda', 'OfferendaV', 'Confractorium', 'Transitorium', 
            // mozarabica
            'Officium', 'OfficiumV', 'Psallendo', 'PsallendoV', 'Tractus', 'TractusV', 
            'Lauda', 'LaudaV', 'Sacrificium', 'SacrificiumV', 'Ad confractionem panis', 
            'Ad accedentes', 'Ad accedentesV', 'Communio'
        ]);

        assert.deepEqual(system3.getValidValues('genre', {type: 'MISS', layer: 'S/C'}), [
            // romana
            'Or', 'Coll', 'Secr', 'Postcomm', 'Superpop', 'VD', 'Infracan', 'BenP',
            // ambrosiana
            'Litaniae', 'Super populum', 'Super sindonem', 'Super oblatam', 'Praefatio', 'Post communionem',
            // mozarabica
            'Oratio post officium', 'Missa', 'Alia oratio', 'Post nomina', 'Ad pacem', 
            'Illatio', 'Post Sanctus', 'Post Pridie', 'Ad orationem dominicam', 'Benedictio', 'Post communionem',
            // gallicana
            'Praefatio', 'Collectio', 'Post prophetiam', 'Post precem', 'Post Aius', 
            'Ante nomina', 'Super oblata', 'Post nomina', 'Ad pacem', 'Secreta', 'Contestatio', 
            'Super munera', 'Immolatio', 'Post Sanctus', 'Post secreta', 'Post mysterium', 
            'Ante orationem dominicam', 'Post orationem dominicam', 'Benedictio populi', 
            'Post communionem', 'Post Eucharistiam', 'Consummatio missae', 'Consecratio',
        ]);

        assert.deepEqual(system3.getValidValues('genre', {type: 'MISS', layer: 'L'}), [
            // romana
            'Proph', 'ProphTrop', 'Lc', 'Ev', 'EvTrop',
            // ambrosiana
            'Lectio', 'Epistola', 'Evangelium',
            // mozarabica
            'Prophetia', 'Sapientia', 'Lectio', 'Epistola', 'Evangelium',
            // gallicana
            'Lectio', 'Epistola', 'Evangelium',
        ]);


        assert.deepEqual(system3.getValidValues('genre', {type: 'OFF', layer: 'G/A'}), [
            'Inv', 'Ant', 'AntV', 'AntB', 'AntQ', 'AntM', 'AntN', 'R', 'RV', 'RTrop', 
            'Rb', 'RbV', 'H', 'Ps', 'CantVT', 'CantNT', 'W', 'WSac',
        ]);

        assert.deepEqual(system3.getValidValues('genre', {type: 'OFF', layer: 'S/C'}), [
            'Cap', 'Or',
        ]);

        assert.deepEqual(system3.getValidValues('genre', {type: 'OFF', layer: 'L'}), [
            'Lec', 'Ser', 'Leg', 'Ev', 'Hom',
        ]);


        assert.deepEqual(system3.getValidValues('genre', {type: 'RIT', layer: 'G/A'}), [
            // romana
            'Intr', 'IntrV', 'IntrTrop', 'Tr', 'Gr', 'GrV', 'H', 'All', 'AllV', 'Seq', 'Off', 'OffV',
            'OffTrop', 'Comm', 'CommV', 'CommTrop',
            // ambrosiana
            'Ingressa ', 'Psalmellus', 'PsalmellusV', 'Cantus', 'CantusV', 'V. in Alleluia',
            'Post evangelium', 'Offerenda', 'OfferendaV', 'Confractorium', 'Transitorium',
            // mozarabica
            'Officium', 'OfficiumV', 'Psallendo', 'PsallendoV', 'Tractus', 'TractusV',
            'Lauda', 'LaudaV', 'Sacrificium', 'SacrificiumV', 'Ad confractionem panis',
            'Ad accedentes', 'Ad accedentesV', 'Communio',

            'Inv', 'Ant', 'AntV', 'AntB', 'AntQ', 'AntM', 'AntN', 'R', 'RV', 'RTrop',
            'Rb', 'RbV', 'H', 'Ps', 'CantVT', 'CantNT', 'W', 'WSac',

        ]);

        assert.deepEqual(system3.getValidValues('genre', {type: 'RIT', layer: 'S/C'}), [
            'F', 'Ex', 'Pf', 'Alloc', 'Lit', 'Abs', 'Ben',
        ]);

        assert.deepEqual(system3.getValidValues('genre', {type: 'RIT', layer: 'L'}), [
            // romana
            'Proph', 'ProphTrop', 'Lc', 'Ev', 'EvTrop',
            // ambrosiana
            'Lectio', 'Epistola', 'Evangelium',
            // mozarabica
            'Prophetia', 'Sapientia', 'Lectio', 'Epistola', 'Evangelium',
            // gallicana
            'Lectio', 'Epistola', 'Evangelium',

            'Lec', 'Ser', 'Leg', 'Ev', 'Hom',
        ]);

        assert.deepEqual(system3.getValidValues('series'), [
            '', 'P', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14'
        ]);
    });

    it('tests System3.validateDigitalPageNumber', function() {
        let row = [];

        row[20] = '20'
        assert(system3.validateDigitalPageNumber(row) === true)
        
        row[20] = ''
        assert(system3.validateDigitalPageNumber(row).constructor.name === 'Error')

        row[20] = '20z'
        assert(system3.validateDigitalPageNumber(row).constructor.name === 'Error')
    });

    it('tests System3.validateGenre', function() {
        let row = [];

        row[17] = 'Intr'
        assert(system3.validateGenre(row))

        row[17] = 'Proph'
        assert(system3.validateGenre(row))

        row[17] = 'Coll'
        assert(system3.validateGenre(row))

        row[17] = 'Ant'
        assert(system3.validateGenre(row))

        row[17] = 'Ser'
        assert(system3.validateGenre(row))

        row[17] = 'Or'
        assert(system3.validateGenre(row))

        row[17] = 'Alloc'
        assert(system3.validateGenre(row))

        row[17] = 'Rub'
        assert(system3.validateGenre(row))

        row[17] = ''
        assert(system3.validateGenre(row).constructor.name === 'Error')

        row[17] = 'invalid'
        assert(system3.validateGenre(row).constructor.name === 'Error')
        
    });

    it('tests System3.validateCeremony', function() {
        let row = [];
        
        system3.indexLabels = [
            {
                name: 'valid label'
            }
        ]

        row[3] = 'MISS'
        row[12] = 'Mass Propers'
        assert(system3.validateCeremony(row))

        row[3] = 'OFF'
        row[12] = 'Office Propers'
        assert(system3.validateCeremony(row))

        row[3] = 'RIT'
        row[12] = 'valid label'
        assert(system3.validateCeremony(row))


        row[3] = 'MISS'
        row[12] = ''
        assert(system3.validateCeremony(row).constructor.name === 'Error')

        row[3] = 'OFF'
        row[12] = 'Office Proper'
        assert(system3.validateCeremony(row).constructor.name === 'Error')

        row[3] = 'RIT'
        row[12] = 'invalid label'
        assert(system3.validateCeremony(row).constructor.name === 'Error')
    });
    
    it('tests System3.validateMassHour', () => {
        let row = [];
        
        row[3] = 'MISS'
        row[11] = 'M2'
        assert(system3.validateMassHour(row))

        row[3] = 'OFF'
        row[11] = 'N1'
        assert(system3.validateMassHour(row))

        row[3] = 'RIT'
        row[11] = ''
        assert(system3.validateMassHour(row))
        
        row[3] = 'MISS'
        row[11] = 'invalid'
        assert(system3.validateMassHour(row).constructor.name === 'Error')
        
        row[3] = 'MISS'
        row[11] = ''
        assert(system3.validateMassHour(row).constructor.name === 'Error')

        row[3] = 'OFF'
        row[11] = 'invalid'
        assert(system3.validateMassHour(row).constructor.name === 'Error')

        row[3] = 'OFF'
        row[11] = ''
        assert(system3.validateMassHour(row).constructor.name === 'Error')

        row[3] = 'RIT'
        row[11] = 'invalid'
        assert(system3.validateMassHour(row).constructor.name === 'Error')
        
    })

    it('tests System3.fillLayer', () => {
        let row = []
        row[3] = 'MISS'
        row[17] = 'All'
        assert.deepEqual(system3.fillLayer(row), [, , , 'MISS', , , , , , , , , , , , , 'G/A', 'All'])
        
        row = []
        row[3] = 'MISS'
        row[17] = 'Proph'
        assert.deepEqual(system3.fillLayer(row), [, , , 'MISS', , , , , , , , , , , , , 'L', 'Proph'])
        
        row = []
        row[3] = 'MISS'
        row[17] = 'Coll'
        assert.deepEqual(system3.fillLayer(row), [, , , 'MISS', , , , , , , , , , , , , 'S/C', 'Coll'])
        
        row = []
        row[3] = 'OFF'
        row[17] = 'Ant'
        assert.deepEqual(system3.fillLayer(row), [, , , 'OFF', , , , , , , , , , , , , 'G/A', 'Ant'])
        
        row = []
        row[3] = 'RIT'
        row[17] = 'Ex'
        assert.deepEqual(system3.fillLayer(row), [, , , 'RIT', , , , , , , , , , , , , 'S/C', 'Ex'])
    })
        
    it('tests System3.fillCeremony', () => {
        let row = []
        
        row[3] = 'MISS'
        assert.deepEqual(system3.fillCeremony(row), [, , , 'MISS', , , , , , , , ,'Mass Propers'])

        row[3] = 'OFF'
        assert.deepEqual(system3.fillCeremony(row), [, , , 'OFF', , , , , , , , ,'Office Propers'])

        row[3] = 'RIT'
        row[12] = 'rit topics'
        assert.deepEqual(system3.fillCeremony(row), [, , , 'RIT', , , , , , , , ,'rit topics'])
    })
    
    it('tests System3.fillPart', () => {
        let row = []
        system3.topics = [
            {
                name: 'commune topic name',
                kind: 1, // C
            },
            {
                name: 'votive topic name',
                kind: null,
                votive: true, // V
            }
        ]
        
        row = []
        row[6] = 'W1'
        assert.deepEqual(system3.fillPart(row), [, , , , 'T', , 'W1'])

        row[5] = 'season'
        row[6] = ''
        row[7] = 'day'
        assert.deepEqual(system3.fillPart(row), [, , , , 'S', 'season', '', 'day'])

        row[5] = ''
        row[6] = ''
        row[7] = ''
        row[10] = 'commune topic name'
        assert.deepEqual(system3.fillPart(row), [, , , , 'C', '', '', '', , , 'commune topic name'])

        row[5] = ''
        row[6] = ''
        row[7] = ''
        row[10] = 'votive topic name'
        assert.deepEqual(system3.fillPart(row), [, , , , 'V', '', '', '', , , 'votive topic name'])
    })
    
    it('tests System3.fillMassHour', () => {
        let row = []

        row[3] = 'MISS'
        row[11] = ''
        assert.deepEqual(system3.fillMassHour(row), [, , , 'MISS', , , , , , , , 'M2'])
        
        row = []
        row[3] = 'MISS'
        row[11] = 'M1'
        assert.deepEqual(system3.fillMassHour(row), [, , , 'MISS', , , , , , , , 'M1'])
        
        row = []
        row[3] = 'OFF'
        row[11] = 'III'
        assert.deepEqual(system3.fillMassHour(row), [, , , 'OFF', , , , , , , , 'III'])
        
    })
    
    it('tests System3.fillSequence', () => {
        let rows = [
            [, , , 'RIT', , , , , , , , , 'ceremony', , ''],
            [, , , 'RIT', , , , , , , , , 'ceremony', , ''],
            [, , , 'RIT', , , , , , , , , 'ceremony', , ''],
            [, , , 'RIT', , , , , , , , , 'ceremony', , ''],
            [, , , 'RIT', , , , , , , , , 'ceremony', , ''],
            [, , , 'RIT', , , , , , , , , 'ceremony', , ''],
            [, , , 'RIT', , , , , , , , , 'ceremony', , ''],
        ]
        
        let rowsResult = [
            [, , , 'RIT', , , , , , , , , 'ceremony', , 1],
            [, , , 'RIT', , , , , , , , , 'ceremony', , 2],
            [, , , 'RIT', , , , , , , , , 'ceremony', , 3],
            [, , , 'RIT', , , , , , , , , 'ceremony', , 4],
            [, , , 'RIT', , , , , , , , , 'ceremony', , 5],
            [, , , 'RIT', , , , , , , , , 'ceremony', , 6],
            [, , , 'RIT', , , , , , , , , 'ceremony', , 7],
        ]
        
        assert.deepEqual(system3.fillSequence(rows), rowsResult)

        rows = [
            [, , , 'RIT', , , , , , , , , 'ceremony', , ''],
            [, , , 'RIT', , , , , , , , , 'ceremony', , ''],
            [, , , 'RIT', , , , , , , , , 'ceremony', , ''],
            [, , , 'MISS', , , , , , , , 'M2', '', , ''],
            [, , , 'MISS', , , , , , , , 'M2', '', , ''],
            [, , , 'MISS', , , , , , , , 'M2', '', , ''],
            [, , , 'MISS', , , , , , , , 'M2', '', , ''],
        ]
        
        rowsResult = [
            [, , , 'RIT', , , , , , , , , 'ceremony', , 1],
            [, , , 'RIT', , , , , , , , , 'ceremony', , 2],
            [, , , 'RIT', , , , , , , , , 'ceremony', , 3],
            [, , , 'MISS', , , , , , , , 'M2', '', , 1],
            [, , , 'MISS', , , , , , , , 'M2', '', , 2],
            [, , , 'MISS', , , , , , , , 'M2', '', , 3],
            [, , , 'MISS', , , , , , , , 'M2', '', , 4],
        ]
        
        assert.deepEqual(system3.fillSequence(rows), rowsResult)
    })


    it('tests System3.fillEmptyFields', () => {
        let rawString = `	748		RIT									minor blessing of water			Incipit exorcismus salis	S/C	F		Adiutorium nostrum	21				KK
	748		RIT									minor blessing of water				S/C	Ex		Exorciso te creatura salis ... immundus adiuratiis.	21				KK
	748		RIT									minor blessing of water				S/C	Or		Immensam clementiam tuam ... spiritalis nequitiae.	21				KK
	748		RIT									minor blessing of water			Sequitur exorcismus aquae	S/C	Ex		Exorciso te creatura aquae ... angelis suis apostaticis	21				KK
	748		RIT									minor blessing of water				S/C	Or		Deus qui ad salutem humani generis ... sit defensa.	21				KK
	748		RIT									minor blessing of water			Hic mittatur sal in aquam.	S/C	F		Fiat haec commixtio salis et aquae	22				KK
	748		RIT									minor blessing of water			Benedictio amborum	S/C	Or		Deus invictae virtutis auctor ... adesse dignetur.	22				KK
	748		RIT									minor blessing of water				S/C	Ben		Benedictio Dei Patris omnipotentis et Filii et Spiritus Sancti descendat super hanc creaturam salis et aquae.Amen	22				KK
	748		RIT									minor blessing of water			Responsorium sine verso dicitur	G/A	Ant		Asperges me Domine	22				KK
	748		RIT									minor blessing of water				S/C	Or		Omnipotens sempiterne Deus qui sacerdotibus tuis ... angeli pacis ingressus.	22				KK`
        let rawData = rawString.split("\n").map(r => r.split("\t"))
        
        let rawResultString = `	748		RIT									minor blessing of water		1	Incipit exorcismus salis	S/C	F		Adiutorium nostrum	21				KK
	748		RIT									minor blessing of water		2		S/C	Ex		Exorciso te creatura salis ... immundus adiuratiis.	21				KK
	748		RIT									minor blessing of water		3		S/C	Or		Immensam clementiam tuam ... spiritalis nequitiae.	21				KK
	748		RIT									minor blessing of water		4	Sequitur exorcismus aquae	S/C	Ex		Exorciso te creatura aquae ... angelis suis apostaticis	21				KK
	748		RIT									minor blessing of water		5		S/C	Or		Deus qui ad salutem humani generis ... sit defensa.	21				KK
	748		RIT									minor blessing of water		6	Hic mittatur sal in aquam.	S/C	F		Fiat haec commixtio salis et aquae	22				KK
	748		RIT									minor blessing of water		7	Benedictio amborum	S/C	Or		Deus invictae virtutis auctor ... adesse dignetur.	22				KK
	748		RIT									minor blessing of water		8		S/C	Ben		Benedictio Dei Patris omnipotentis et Filii et Spiritus Sancti descendat super hanc creaturam salis et aquae.Amen	22				KK
	748		RIT									minor blessing of water		9	Responsorium sine verso dicitur	G/A	Ant		Asperges me Domine	22				KK
	748		RIT									minor blessing of water		10		S/C	Or		Omnipotens sempiterne Deus qui sacerdotibus tuis ... angeli pacis ingressus.	22				KK`
        let resultData = rawResultString.split("\n").map(r => r.split("\t"))
        
        let sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getDataRows').atLeast(1).returns(rawData)
        sheetWrapperMock.expects('setValuesAt').atLeast(1).withArgs(resultData, {row: 2, col: 1, rows: 10, cols: 25})
        
        system3.fillEmptyFields(false)
        assert.deepEqual(system3.data, resultData)
        
        sheetWrapperMock.verify()
        sheetWrapperMock.restore()
    })
    
    it('tests System3.migrateGenre', () => {
        let genres = [
            ['F'],
            ['Ex'],
            ['Or1'],
            ['Ex'],
            ['Or2'],
            [null],
            ['Or'],
            ['Ben'],
            ['Ant'],
            ['Or3'],
        ];
        
        let expected = [
            ['F'],
            ['Ex'],
            ['Or'],
            ['Ex'],
            ['Or'],
            [null],
            ['Or'],
            ['Ben'],
            ['Ant'],
            ['Or'],
        ];
        
        let sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getValuesAtColumn').atLeast(1).returns(genres)
        sheetWrapperMock.expects('setValuesAt').atLeast(1).withArgs(expected, {row: 2, col: 15, rows: 10, cols: 1})
        
        system3.migrateGenre()
        
        sheetWrapperMock.verify()
        sheetWrapperMock.restore()
    })
    
    it('tests Gen is empty', () => {
        let rawString = '	748		MISS	T	Qu	Gen	Gen		Corp	Corp	M2	Mass Propers		1	Officium de Corpore Christi a LXX usque ad Pascha.	G/A	Intr		Panem caeli dedit nobis	376	151v			KK'
        let rawData = rawString.split("\n").map(r => r.split("\t"))

        let rawResultString = '	748		MISS	V	Qu	Gen	Gen		Corp	Corp	M2	Mass Propers		1	Officium de Corpore Christi a LXX usque ad Pascha.	G/A	Intr		Panem caeli dedit nobis	376	151v			KK'
        let resultData = rawResultString.split("\n").map(r => r.split("\t"))
        
        let sheetWrapperMock = sinon.mock(system3.sheetWrapper)
        sheetWrapperMock.expects('getDataRows').atLeast(1).returns(rawData)
        sheetWrapperMock.expects('setValuesAt').atLeast(1).withArgs(resultData, {row: 2, col: 1, rows: 1, cols: 25})
        
        system3.fillEmptyFields(false)
        assert.deepEqual(system3.data, resultData)
        
        sheetWrapperMock.verify()
        sheetWrapperMock.restore()
        
    })
});

