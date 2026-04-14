using System;
using System.IO.Packaging;

namespace OpenXmlStreaming;

/// <summary>
/// Describes a relationship to be added to a part when using <see cref="OpenXmlPackageWriter.WritePart"/>.
/// </summary>
public readonly struct PartRelationship
{
    public PartRelationship(Uri targetUri, string relationshipType, TargetMode targetMode = TargetMode.Internal, string? id = null)
    {
        TargetUri = targetUri;
        RelationshipType = relationshipType;
        TargetMode = targetMode;
        Id = id;
    }

    public Uri TargetUri { get; }

    public string RelationshipType { get; }

    public TargetMode TargetMode { get; }

    public string? Id { get; }
}
