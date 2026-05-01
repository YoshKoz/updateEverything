@{
    Rules        = @{
        PSAvoidAssignmentToAutomaticVariable = @{
            Enable = $false
        }
    }
    ExcludeRules = @(
        'PSAvoidAssignmentToAutomaticVariable'
        'PSAvoidUsingWriteHost'
        'PSAvoidUsingEmptyCatchBlock'
        'PSUseShouldProcessForStateChangingFunctions'
        'PSUseSingularNouns'
        'PSAvoidUsingPositionalParameters'
        'PSUseBOMForUnicodeEncodedFile'
        'PSUseApprovedVerbs'
    )
}
