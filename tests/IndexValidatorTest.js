const sinon = require('sinon')
const assert = require('assert')
const path = require('path')
const NodeVMHelper = require('./NodeVMHelper.js')

let indexValidator
let indexValidatorMock

before(async () => {
    let vmHelper = new NodeVMHelper()
    await vmHelper.loadContext(path.resolve(__dirname, '../src'))
    indexValidator = vmHelper.instantiateClass('IndexValidator')
    
    indexValidatorMock = sinon.mock(indexValidator)

    indexValidatorMock.expects('getBooksData').atLeast(1).returns([
        {
            id: 1,
            number_of_pages: 123
        },
        {
            id: 2,
            number_of_pages: 123
        },
        {
            id: 3,
            number_of_pages: 123
        },
        {
            id: 4,
            number_of_pages: 123
        },
        {
            id: 785,
            number_of_pages: 123
        }
    ])

    indexValidator.loadBooksData()

    indexValidatorMock.expects('getIndexLabelData').atLeast(1).returns([
        {
            "id": 1,
            "name": "valid"
        },
        {
            "id": 2,
            "name": "Temporal Masses"
        }
    ])
    indexValidator.loadIndexLabelData()
});
 
describe('tests IndexValidator', function() {
    it('tests IndexValidator.isValidBookIdValue', function() {
        assert.equal(indexValidator.isValidBookIdValue(''), false)
        assert.equal(indexValidator.isValidBookIdValue(null), false)
        assert.equal(indexValidator.isValidBookIdValue(undefined), false)
        assert.equal(indexValidator.isValidBookIdValue('abc'), false)
        assert.equal(indexValidator.isValidBookIdValue(1234), true)
        assert.equal(indexValidator.isValidBookIdValue('1234'), true)
        assert.equal(indexValidator.isValidBookIdValue('1234 '), true)
        assert.equal(indexValidator.isValidBookIdValue(' 1234'), true)
    });

    it('tests IndexValidator.isValidBookId', function() {
        assert.equal(indexValidator.isValidBookId('1'), true)
        assert.equal(indexValidator.isValidBookId('2'), true)
        assert.equal(indexValidator.isValidBookId('5'), false)

        assert.equal(indexValidator.isValidBookId(1), true)
        assert.equal(indexValidator.isValidBookId(2), true)
        assert.equal(indexValidator.isValidBookId(5), false)

        assert.equal(indexValidator.isValidBookId('1 '), true)
        assert.equal(indexValidator.isValidBookId(' 1'), true)
    });

    it('tests IndexValidator.isValidDigitalPageNumberValue', function() {
        assert.equal(indexValidator.isValidDigitalPageNumberValue(''), false)
        assert.equal(indexValidator.isValidDigitalPageNumberValue(null), false)
        assert.equal(indexValidator.isValidDigitalPageNumberValue(undefined), false)
        assert.equal(indexValidator.isValidDigitalPageNumberValue('abc'), false)
        assert.equal(indexValidator.isValidDigitalPageNumberValue(1234), true)
        assert.equal(indexValidator.isValidDigitalPageNumberValue('1234'), true)
        assert.equal(indexValidator.isValidDigitalPageNumberValue('1234 '), true)
        assert.equal(indexValidator.isValidDigitalPageNumberValue(' 1234'), true)
    });


    it('tests IndexValidator.isValidDigitalPageNumber', function() {
        assert.equal(indexValidator.isValidDigitalPageNumber('1', '1'), true)
        assert.equal(indexValidator.isValidDigitalPageNumber('999', '1'), false)
        assert.equal(indexValidator.isValidDigitalPageNumber('1', '99'), false)
        assert.equal(indexValidator.isValidDigitalPageNumber('999', '99'), false)

        assert.equal(indexValidator.isValidDigitalPageNumber(1, 1), true)
        assert.equal(indexValidator.isValidDigitalPageNumber(999, 1), false)
        assert.equal(indexValidator.isValidDigitalPageNumber(1, 99), false)
        assert.equal(indexValidator.isValidDigitalPageNumber(999, 99), false)
    });

    it('tests IndexValidator.splitIndexLabels', function() {
        assert.deepEqual(indexValidator.splitIndexLabels(''), [''])
        assert.deepEqual(indexValidator.splitIndexLabels('abc'), ['abc'])
        assert.deepEqual(indexValidator.splitIndexLabels('abc, cba'), ['abc', 'cba'])
        assert.deepEqual(indexValidator.splitIndexLabels('abc, cba '), ['abc', 'cba'])
        assert.deepEqual(indexValidator.splitIndexLabels(' abc, cba '), ['abc', 'cba'])
        assert.deepEqual(indexValidator.splitIndexLabels(' abc, cba , zzz'), ['abc', 'cba', 'zzz'])

        assert.deepEqual(indexValidator.splitIndexLabels(1234), ['1234'])
        assert.deepEqual(indexValidator.splitIndexLabels(null), [])
        assert.deepEqual(indexValidator.splitIndexLabels(undefined), [])
    });

    it('tests IndexValidator.isValidIndexLabel', function() {
        assert.equal(indexValidator.isValidIndexLabel('valid'), true)
        assert.equal(indexValidator.isValidIndexLabel('In Valid'), false)
    });

    it('tests IndexValidator.validate with missing rows', function() {

        indexValidatorMock.expects('getSheetData').atLeast(1).returns([])
        assert.deepEqual(indexValidator.validate(), [
            {
                row: 1,
                col: 1,
                message: 'Missing rows?'
            }
        ])

    })

    it('tests IndexValidator.validate with full valid data', function() {
        indexValidatorMock.expects('getSheetData').atLeast(1).returns([
            [
                '785',
                '',
                '1',
                '',
                '',
                'Temporal Masses',
                '',
                '',
                '',
                '',
                'HB',
                ''
            ]
        ])

        assert.deepEqual(indexValidator.validate(), [])
    })

    it('tests IndexValidator.validate with missing book', function() {
        indexValidatorMock.expects('getSheetData').atLeast(1).returns([
            [
                '999',
                '',
                '1',
                '',
                '',
                'Temporal Masses',
                '',
                '',
                '',
                '',
                'HB',
                ''
            ]
        ])

        assert.deepEqual(indexValidator.validate(), [
            {
                row: 2,
                col: 1,
                message: 'Unknown book id: 999'
            },
            {
                row: 2,
                col: 3,
                message: 'Invalid digital page number in book: 999'
            }
        ])
    })

    it('tests IndexValidator.validate with missing digital page number', function() {
        indexValidatorMock.expects('getSheetData').atLeast(1).returns([
            [
                '1',
                '',
                '999',
                '',
                '',
                'Temporal Masses',
                '',
                '',
                '',
                '',
                'HB',
                ''
            ]
        ])

        assert.deepEqual(indexValidator.validate(), [
            {
                row: 2,
                col: 3,
                message: 'Invalid digital page number in book: 1'
            }
        ])
    })

    it('tests IndexValidator.validate with invalid index label', function() {
        indexValidatorMock.expects('getSheetData').atLeast(1).returns([
            [
                '1',
                '',
                '1',
                '',
                '',
                'invalid index label',
                '',
                '',
                '',
                '',
                'HB',
                ''
            ]
        ])

        assert.deepEqual(indexValidator.validate(), [
            {
                row: 2,
                col: 6,
                message: 'Unknown index label: invalid index label'
            }
        ])
    })

    it('tests IndexValidator.validate with partial invalid index label', function() {
        indexValidatorMock.expects('getSheetData').atLeast(1).returns([
            [
                '1',
                '',
                '1',
                '',
                '',
                'Temporal Masses, invalid index label',
                '',
                '',
                '',
                '',
                'HB',
                ''
            ]
        ])

        assert.deepEqual(indexValidator.validate(), [
            {
                row: 2,
                col: 6,
                message: 'Unknown index label: invalid index label'
            }
        ])
    })

    it('tests IndexValidator.validate with partial invalid index label with case insensitive match', function() {
        indexValidatorMock.expects('setIndexLabel').once().withArgs('Temporal Masses, invalid index label', 2)
        indexValidatorMock.expects('getSheetData').atLeast(1).returns([
            [
                '1',
                '',
                '1',
                '',
                '',
                'temporal masses, invalid index label',
                '',
                '',
                '',
                '',
                'HB',
                ''
            ]
        ])

        assert.deepEqual(indexValidator.validate(), [
            {
                row: 2,
                col: 6,
                message: 'Unknown index label: invalid index label'
            }
        ])
    })

    it('tests IndexValidator.validate with partial case insensitive match', function() {
        indexValidatorMock.expects('setIndexLabel').once().withArgs('Temporal Masses, Temporal Masses', 2)
        indexValidatorMock.expects('getSheetData').atLeast(1).returns([
            [
                '1',
                '',
                '1',
                '',
                '',
                'temporal masses, Temporal Masses',
                '',
                '',
                '',
                '',
                'HB',
                ''
            ]
        ])

        assert.deepEqual(indexValidator.validate(), [])
    })

    it('tests IndexValidator.validate with case insensitive match', function() {
        indexValidatorMock.expects('setIndexLabel').once().withArgs('Temporal Masses', 2)
        indexValidatorMock.expects('getSheetData').atLeast(1).returns([
            [
                '1',
                '',
                '1',
                '',
                '',
                'temporal masses',
                '',
                '',
                '',
                '',
                'HB',
                ''
            ]
        ])

        assert.deepEqual(indexValidator.validate(), [])
    })

});
